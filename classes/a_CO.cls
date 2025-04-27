VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mPay_Ex As Long
Dim mVAT As Long
Dim mDisc As Long
Dim mPay As Long
Dim mPayExtras As Long
Dim lngLDeposit_tot As Long
Dim mServiceItem As Long
Private oSM As z_StockManager
Private oFSO As FileSystemObject

Dim lngTotal1 As Long
Dim lngTotal2 As Long
Dim lngTotal3 As Long
Dim lngTotal4 As Long
Dim lngQty_tot As Long
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
Dim iSTartBillToIdx As Integer
Dim iSTartDellToIdx As Integer
Event Valid(pMsg As String)
Event TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Private mudtProps As COrderProps
Private mcolStack As Collection

Private WithEvents mColCOLines As ch_COL
Attribute mColCOLines.VB_VarHelpID = -1
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private colClassors As Collection
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
Private oCurrency_Foreign As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Const CDOCCODE = "CO"
Private xMLDoc As ujXML
Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Private Sub mColCOLines_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.mColCOLines_CurrRowStatus(pMsg)", pMsg
End Sub
Public Property Get ForeignCurrency() As a_Currency
    Set ForeignCurrency = oCurrency_Foreign
End Property
Private Sub mColCOLines_RowsChange()
    On Error GoTo errHandler
'occurs for new rows and deletion of rows
    
    CalculateTotal
    RaiseEvent TotalChange(lngTotal2, Format(lngTotal2 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
                        lngTotalDeposit, Format(lngTotalDeposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
                        lngTotalVAT, Format(lngTotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString))
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.mColCOLines_RowsChange"
End Sub
Public Function CalculateTotal()
    On Error GoTo errHandler
Dim oCOL As a_COL
Dim bLineDiscountApplied As Boolean
    mobjValid.RuleBroken "OVERFLOW", False
    bLineDiscountApplied = False
    mPay = 0
    lngLDeposit_tot = 0
    mVAT = 0
    mDisc = 0
    mPay_Ex = 0
    mServiceItem = 0
    lngQty_tot = 0
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    iExtrasCount = 0
    For Each oCOL In mColCOLines
        If oCOL.Fulfilled <> "CAN" Then
            oCOL.CalculateLine
            If oCOL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            If oCOL.ServiceItem = False Then
                lngQty_tot = lngQty_tot + oCOL.Qty
            End If
'
            If oCOL.ServiceItem = True Then
                mServiceItem = mServiceItem + oCOL.Ext
            End If
            mPay_Ex = mPay_Ex + oCOL.Ext_Ex + oCOL.ExtraCharge
            mVAT = mVAT + oCOL.VAT + (oCOL.ExtraCharge - (CDbl(oCOL.ExtraCharge) * 100 / (100 + oCOL.ExtraVATRate)))
            mDisc = mDisc + oCOL.Disc
            mPay = mPay + oCOL.Ext + oCOL.ExtraCharge
            If oCOL.BottomOfDocument Then
                mPayExtras = mPayExtras + oCOL.Ext
            End If
        End If
    Next
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtras = mPayExtras
    mudtProps.TotalDiscount = mDisc
    mudtProps.TotalNonVAT = mPay_Ex
    mudtProps.TotalVAT = mVAT
    mudtProps.TotalExtension = mPay
    mudtProps.TotalPayable = mPay '- lngLDeposit_tot
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
    ErrorIn "a_CO.CalculateTotal"
End Function
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ShowVAT"
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As COrderData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As COrderData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "COLS", mColCOLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.GetSuperState"
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
    mColCOLines.SetSuperState .ReadProperty("COLS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
    mColCOLines.BeginEdit
  ' save object state
    mcolStack.Add GetState
 ' Debug.Print "BeginEdit " & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
    mColCOLines.CancelEdit
    mudtProps.IsDeleted = False
    ' restore object state
    With mcolStack
      SetState .Item(.Count)
      .Remove .Count
    End With
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.CancelEdit"
End Sub

Public Sub ApplyEdit(pResult As String)
    On Error GoTo errHandler
Dim objPersist As a_CO_P
Dim strBuffer As String
Dim oCOL As a_COL

    For Each oCOL In Me.COLines
        If oCOL.Ref = "" And oCOL.ServiceItem = False Then oCOL.SetRef Me.OrderRef
    Next
 '   CalculateTotal
    mudtProps.TPID = Me.Customer.ID
    'Set these values fromthe objects before saving -  the objects cant be saved
    If Not (oBillToAddress Is Nothing) Then
        mudtProps.BillToAddressID = oBillToAddress.ID
    End If
    If Not (oDelToAddress Is Nothing) Then
        mudtProps.GoodsAddressID = oDelToAddress.ID
    End If
    If DOCCode = "" Then
        DOCCode = GetCONumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
    
    If mcolStack.Count = 0 Then Err.Raise 445
 '   oPC.COShort.BeginTrans
    mColCOLines.ApplyEdit
    Set objPersist = New a_CO_P
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.TRID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If mudtProps.IsNew Then mudtProps.CaptureDate = Now()
        If Not IsValid Then Err.Raise 445
        mcolStack.Remove mcolStack.Count
        strBuffer = objPersist.Save(GetSuperState)
        If strBuffer > "" Then SetSuperState strBuffer
        mudtProps.IsNew = False
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
 '   oPC.COShort.CommitTrans
    mudtProps.IsDirty = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ApplyEdit(pResult)", pResult
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mColCOLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IsValid"
End Property
Public Function GetStatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.GetSTatus"
End Function
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mColCOLines = New ch_COL
    mColCOLines.component Me
    Set mCustomer = New a_Customer

    Set oPCL = New z_PrintClient
    mudtProps.IsNew = True
    mOriginalStatus = stInProcess
    Set oBillToAddress = New a_Address
    Set oDelToAddress = New a_Address
    mobjValid.BreakRule "TP", True
    LoadClassorsCollection
    mudtProps.DOCCode = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
      
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mColCOLines = Nothing
    Set mCustomer = Nothing
    Set colClassors = Nothing
    Set oPCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_CO_P
Dim oAdd As a_Address
Dim oComp As a_Company
Dim oCurr As a_Currency
Dim i As Integer

    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_CO_P
    SetSuperState objPersist.Fetch(TRID)
    mOriginalStatus = Me.Status
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
    If GoodsAddressID > 0 Then
        Set oDelToAddress = Customer.Addresses.FindByID(GoodsAddressID)
        'The address might not be found if it has been removed or the customer has been merged with another
        If oDelToAddress Is Nothing Then
            If Not Customer.DelToAddress Is Nothing Then
                Set oDelToAddress = Customer.DelToAddress
                GoodsAddressID = Customer.DelToAddress.ID
            End If
        End If
    End If
    
    Set objPersist = Nothing
    mobjValid.RuleBroken "TP", False

    mColCOLines_RowsChange
    Me.GetStatus
    iSTartDellToIdx = 0
    If Not oDelToAddress Is Nothing Then
    For i = 1 To Me.Customer.Addresses.Count
        If Customer.Addresses(i).ID = oDelToAddress.ID Then
            iSTartDellToIdx = i
        End If
    Next
    End If
    iSTartBillToIdx = 0
    If Not oBillToAddress Is Nothing Then
    For i = 1 To Me.Customer.Addresses.Count
        If Customer.Addresses(i).ID = oBillToAddress.ID Then
            iSTartBillToIdx = i
        End If
    Next
    End If
    SetDirty False
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub
Public Function LoadWantsForTP(TPID As Long) As Boolean
    On Error GoTo errHandler
Dim objPersist As a_CO_P
Dim oAdd As a_Address
Dim oComp As a_Company
Dim oCurr As a_Currency
Dim tmpBuffer As String

    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_CO_P
    tmpBuffer = objPersist.FetchWants(TPID)
    If tmpBuffer > " " Then
        SetSuperState tmpBuffer
        mOriginalStatus = Me.Status
        mCustomer.Load mudtProps.TPID
        For Each oAdd In Customer.Addresses
            If oAdd.ID = mudtProps.BillToAddressID Then
                Set oBillToAddress = oAdd
            End If
            If oAdd.ID = mudtProps.GoodsAddressID Then
                Set oDelToAddress = oAdd
            End If
        Next
        mobjValid.RuleBroken "TP", False
        mColCOLines_RowsChange
        Me.GetStatus
        LoadWantsForTP = True
    Else
        LoadWantsForTP = False
    End If
    Set objPersist = Nothing
'H:
'    MsgBox "a_CO:Load, " & or
'    Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.LoadWantsForTP(TPID)", TPID
End Function

Public Property Get Customer() As a_Customer
    Set Customer = mCustomer
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Delete"
End Sub
Public Property Get Total_2F() As String
    On Error GoTo errHandler
    Total_2F = Format(lngTotal2 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Total_2F"
End Property
Public Property Get Total_3F() As String
    On Error GoTo errHandler
    Total_3F = Format(lngTotal3 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Total_3F"
End Property
Public Property Get Total_4F() As String
    On Error GoTo errHandler
    Total_4F = Format(lngTotal4 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Total_4F"
End Property
Public Property Get TotalQty() As Long
    On Error GoTo errHandler
    TotalQty = lngQty_tot
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalQty"
End Property
Public Property Get TotalQtyF() As String
    On Error GoTo errHandler
    TotalQtyF = Format(lngQty_tot, "#,##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalQtyF"
End Property
Public Property Get TotalDepositF() As String
    On Error GoTo errHandler
    TotalDepositF = Format(lngTotalDeposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalDepositF"
End Property
Public Property Get TotalVATF() As String
    On Error GoTo errHandler
    TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalVATF"
End Property
Public Property Get TotalPayableF() As String
    On Error GoTo errHandler
    TotalPayableF = Format(mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalPayableF"
End Property

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TRID(val)", val
End Property
Public Property Get OrderType() As enumCOType
    On Error GoTo errHandler
    OrderType = mudtProps.OrderType
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.OrderType"
End Property
Public Property Let OrderType(val As enumCOType)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OrderType = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.OrderType(val)", val
End Property
Public Property Get OrderRef() As String
    On Error GoTo errHandler
    OrderRef = FNS(mudtProps.OrderNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.OrderRef"
End Property
Public Property Let OrderRef(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OrderNum = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.OrderRef(val)", val
End Property
Public Property Get BillToAddressID() As Long
    On Error GoTo errHandler
    BillToAddressID = mudtProps.BillToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.BillToAddressID:Get"
End Property
Public Property Let BillToAddressID(val As Long)
    On Error GoTo errHandler
    mudtProps.BillToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.BillToAddressID:Let"
End Property

Public Sub SetBillToAddress(val As a_Address)
Dim i As Integer

    On Error GoTo errHandler
   ' If mcolStack.Count = 0 Then Err.Raise 383
    If val Is Nothing Then Exit Sub
    Set oBillToAddress = val
    mudtProps.BillToAddressID = oBillToAddress.ID
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetBillToAddress(val)", val
End Sub
Public Property Get GoodsAddressID() As Long
    On Error GoTo errHandler
    GoodsAddressID = mudtProps.GoodsAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.GoodsAddressID:Get"
End Property
Public Property Let GoodsAddressID(val As Long)
    On Error GoTo errHandler
    mudtProps.GoodsAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.GoodsAddressID:Let"
End Property
Public Sub setDelToAddress(val As a_Address)
Dim i As Integer
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If val Is Nothing Then Exit Sub
    Set oDelToAddress = val
    mudtProps.GoodsAddressID = oDelToAddress.ID
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.setDelToAddress(val)", val
End Sub
Public Property Get StartPosOfBillToIdx() As Integer
    StartPosOfBillToIdx = iSTartBillToIdx
End Property
Public Property Get StartPosOfDEllToIdx() As Integer
    StartPosOfDEllToIdx = iSTartDellToIdx
End Property

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
'    mColCOLines_RowsChange  'recalulates andd raises event
'    RaiseEvent reloadlist
'End Function
Public Function SetCustomer(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
Dim i As Integer

    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
        Set oBillToAddress = Me.Customer.BillTOAddress
        Set oDelToAddress = Me.Customer.DelToAddress
        
        For i = 1 To Me.Customer.Addresses.Count
            If Customer.Addresses(i).ID = oDelToAddress.ID Then
                iSTartDellToIdx = i
            End If
        Next
        For i = 1 To Me.Customer.Addresses.Count
            If Customer.Addresses(i).ID = oBillToAddress.ID Then
                iSTartBillToIdx = i
            End If
        Next
        
        
        
        SetDirty True
        
        
        
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetCustomer(pTPID)", pTPID
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
    ErrorIn "a_CO.SetCustomerFromAccNum(pAccnum)", pAccnum
End Function
Public Property Get ORGUID() As String
    On Error GoTo errHandler
    ORGUID = mudtProps.ORGUID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ORGUID"
End Property
Public Property Let ORGUID(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ORGUID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ORGUID(val)", val
End Property

Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.StaffID(val)", val
End Property

Public Property Get StaffName() As String
    On Error GoTo errHandler
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.StaffName"
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
    ErrorIn "a_CO.StaffNameB"
End Property
Public Property Get Signature(Optional bSkipBrackets As Boolean) As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
        Signature = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).Signature
    If bSkipBrackets Then
        If strT > "" Then
            Signature = strT
        Else
            Signature = ""
        End If
    Else
        If strT > "" Then
            Signature = " (" & strT & ")"
        Else
            Signature = ""
        End If
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Signature"
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
Public Property Get TPNAME() As String
    On Error GoTo errHandler
    TPNAME = FNS(mudtProps.TPNAME)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPName"
End Property
Public Property Let TPNAME(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPNAME = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPName(val)", val
End Property

Public Property Get TPPhone() As String
    On Error GoTo errHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPPhone"
End Property
Public Property Let TPPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
    On Error GoTo errHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPFax"
End Property
Public Property Let TPFax(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPFax(val)", val
End Property

Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TPACCNum(val)", val
End Property

Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Memo"
End Property
Public Function SetMemo(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
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
Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = FNS(mudtProps.OrderNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Memo"
End Property
Public Function SetRef(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.OrderNum) Then
        Err.Raise 384
    End If
    If val <> mudtProps.OrderNum Then
        mudtProps.OrderNum = val
        SetDirty True
    End If
    SetRef = True
    Exit Function
End Function
Public Property Get ForAttn() As String
    On Error GoTo errHandler
    ForAttn = FNS(mudtProps.ForAttn)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ForAttn"
End Property
Public Function SetForAttn(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
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

Public Property Get BusPhone() As String
    On Error GoTo errHandler
    BusPhone = FNS(mudtProps.BusPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.BusPhone"
End Property
Public Property Let BusPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BusPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.BusPhone(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.DocCode"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.DocCode(val)", val
End Property

'Public Property Get CurrencyFormat() As String
'    CurrencyFormat = FNS(mudtProps.CurrencyFormat)
'    If CurrencyFormat <= "" Then CurrencyFormat = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").FormatString
'End Property
'Public Property Let CurrencyFormat(Val As String)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.CurrencyFormat = Val
'    SetDirty True
'End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.DocDateF"
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
    ErrorIn "a_CO.DocDate(val)", val
End Property

Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IssDate"
End Property
Public Property Get IssDateF() As String
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.CaptureDateFormatted"
End Property
Public Property Let IssDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.IssDate(val)", val
End Property

'Public Property Get ForeignCurrency() As a_Currency
'    Set ForeignCurrency = oCurrency_Foreign
'End Property

'Public Property Get CurrencyFactor() As Double
'    If mudtProps.CurrencyFactor = 0 Then mudtProps.CurrencyFactor = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").Factor
'    CurrencyFactor = mudtProps.CurrencyFactor
'End Property
'Public Property Let CurrencyFactor(Val As Double)
'    mudtProps.CurrencyFactor = Val
'    SetDirty True
'End Property
'
'Public Property Get VATRate() As Double
'    VATRate = mudtProps.VATRate
'End Property
'Public Property Get VATRateFormatted() As String
'    VATRate = Format(mudtProps.VATRate, "##0.00")
'End Property
'Public Property Let VATRate(Val As Double)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.VATRate = Val
'    mudtProps.IsDirty = True
'End Property

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
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.VATable(val)", val
End Property


Public Property Get COLines() As ch_COL
  Set COLines = mColCOLines
End Property
Public Property Get DisplayCO() As d_CO
    On Error GoTo errHandler
    Set DisplayCO = New d_CO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.DisplayCO"
End Property
Public Property Get StatusF() As String
    On Error GoTo errHandler
    If Me.OrderType = enNormalCO Then
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
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.StatusF"
End Property
Public Property Get StatusForPrinting() As String
    On Error GoTo errHandler
    If Me.OrderType = enNormalCO Then
        If mudtProps.Status = stVOID Then
            StatusForPrinting = "VOID"
        ElseIf mudtProps.Status = stInProcess Then
            StatusForPrinting = "IN PROCESS"
        ElseIf mudtProps.Status = stISSUED Then
            StatusForPrinting = ""
        ElseIf mudtProps.Status = stCOMPLETE Then
            StatusForPrinting = "COMPLETE"
        ElseIf mudtProps.Status = stPROFORMA Then
            StatusForPrinting = "PROFORMA"
        ElseIf mudtProps.Status = stCANCELLED Then
            StatusForPrinting = "CANCELLED"
        End If
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.StatusForPrinting"
End Property

Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Status"
End Property
Public Function GetCONumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strWork As String

    Set oCode = New z_Code
    
    strWork = oCode.GetNextCode(enCustomerOrder, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetCONumber = strDefaultCo & "C" & strWork
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.GetCONumber(strTmp)", strTmp
End Function

Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property
Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property
Public Function Print_Display(Optional ByVal pRoundedUp As Boolean) As String
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim oCOL As a_COL
Dim strCO As String
Dim i As Integer

    CalculateTotal
    strCO = "Order: " & Me.DOCCode & vbCrLf
    For i = 1 To Me.COLines.Count
        If COLines(i).BottomOfDocument = False Then
            With Me.COLines(i)
                If .Discount > 0 Then bDiscountExists = True
                strCO = strCO & .CodeF & "  " & .QtyF & "  " & Left(.TitleAuthorPublisher, 30) & "  " & .PriceF & "  " & IIf(.Discount = 0, "", .DiscountF) & "  " & .ExtensionF
            End With
        strCO = strCO & vbCrLf
        Else
            iExtrasCount = iExtrasCount + 1
        End If
    Next i

    strCO = strCO & vbCrLf
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strCO = strCO & "subtotal    " & Total_2F & vbCrLf
'        For i = 1 To COLines.Count
'            If COLines(i).BottomOfDocument = True Then
'                With COLines(i)
'                    If .ServiceItem Then
'                        strCO = strCO & .TitleAuthorPublisher & "    " & .ExtensionF & vbCrLf
'                    End If
'                End With
'            End If
'        Next i
    End If
    If lngTotalDeposit > 0 Then
        strCO = strCO & "Total    " & Total_3F & vbCrLf
        strCO = strCO & "Deposits paid    " & TotalDepositF & vbCrLf
    End If
  '  strCO = strCO & "Total " & TotalPayableF & vbCrLf
  '  strCO = strCO & "Includes VAT of " & TotalVATF & vbCrLf
    Print_Display = strCO
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
    ErrorIn "a_CO.Print_Display(pRoundedUp)", pRoundedUp
End Function
Public Function PrintCO() As Boolean
    On Error GoTo errHandler
Dim iTotalLinesCOunt As Integer
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl

    
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintCO = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("CO_" & Me.DOCCode) = False Then
            PrintCO = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
        End With
        PrintCO = True
    End If
    
    
    
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt COLines.Count
'    iTotalLinesCOunt = 1
    If lngTotalDeposit > 0 Then
        iTotalLinesCOunt = 3
    Else
        iTotalLinesCOunt = 2
    End If
    oPCL.LineTotalCount iTotalLinesCOunt 'This must be written before the TOTAL tags for the Printserver requirements
    oPCL.WriteTag "TOTAL", "Total qty" & "|" & Me.TotalQtyF

    If lngTotalDeposit > 0 Then
        oPCL.WriteTag "TOTAL", "Total" & "|" & Me.Total_3F
        oPCL.WriteTag "TOTAL", "Less deposits paid" & "|" & Me.TotalDepositF
    Else
        oPCL.WriteTag "TOTAL", "Total" & "|" & Me.Total_3F
    End If
    
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "LOGO", oPC.SharedFolderRoot & "\LOGOS\" & oPC.Configuration.DefaultCompany.LogoFilePath
    oPCL.WriteTag "OURADDRESS", oPC.Configuration.DefaultCompany.StreetAddressDelimited
    If Not Me.BillTOAddress Is Nothing Then
        oPCL.WriteTag "PHONE", "Phone: " & Me.Customer.Phone
        oPCL.WriteTag "FAX", IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
        oPCL.WriteTag "BILLTO", Me.BillTOAddress.AddressDelimited
    End If
    If Not Me.DelToAddress Is Nothing Then
        oPCL.WriteTag "DELTO", Me.DelToAddress.AddressDelimited
    End If
    oPCL.WriteTag "COREGNUMBER", oPC.Configuration.DefaultCompany.CoRegistrationNumber
    oPCL.WriteTag "VATNUMBER", oPC.Configuration.DefaultCompany.VatNumber
    oPCL.WriteTag "MEMO", PackText(Me.Memo)
    oPCL.WriteTag "STORECODE", oPC.Configuration.DefaultStore.code
    oPCL.WriteTag "CUSTOMER", Me.Customer.Fullname & IIf(Customer.AcNo > "", " (" & Customer.AcNo & ")", "")
    oPCL.WriteTag "DATE", Me.IssDateF
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    FillTableRows
    oPCL.FinishRequest

EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.PrintCO"
End Function
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
    On Error GoTo errHandler
Dim strLabels As String
Dim strValues As String
Dim i As Integer
    strLabels = ""
    strValues = ""

    strLabels = "Total qty: " & Me.TotalQtyF
    If lngTotalDeposit > 0 Then
        strLabels = strLabels & Chr(13) & "Total"
        strValues = strValues & Chr(13) & Me.TotalPayableF
        strLabels = strLabels & Chr(13) & "Deposits paid   "
        strValues = strValues & Chr(13) & Me.TotalDepositF
    Else
        strLabels = strLabels & Chr(13) & "Total"
        strValues = strValues & Chr(13) & Me.TotalPayableF
    End If
    pLabels = strLabels
    pValues = strValues
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Function FillTableRows()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To COLines.Count
        With COLines(i)
            If oPC.AllowsSSInvoicing Then
                oPCL.COLSend .CodeF, .QtyFirmF & "/" & .QtySSF, .TitleAuthorPublisher, .Ref, .PriceF, .ExtensionF, .DiscountF, .DepositF, .Note '"InvDate", oInv.TransDateFormatted
            Else
                oPCL.COLSend .CodeF, .QtyF, .TitleAuthorPublisher, .Ref, .PriceF, .ExtensionF, .DiscountF, .DepositF, .Note '"InvDate", oInv.TransDateFormatted
            End If
        End With
    Next i
EXIT_Handler:
'ErrHandler:
'    MsgBox error
'    GoTo EXIT_Handler
   ' Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.FillTableRows"
End Function
Public Function ExportToXML(pFilename As String, DispatchMode As enTransmitType, _
            Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, Optional pQtyCopies As Integer, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
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
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim oDC As a_DocumentControl
Dim strExecutable As String
Dim cnt As Integer
Dim strTemplateName As String

    If DispatchMode = enMail Then
        If oPC.UsesOutlookForCOEmail Then
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
                        p 1
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
        ExportToXML = False
        Exit Function
    End If
                        p 2
    If pQtyCopies < 1 Then pQtyCopies = 1

    Set xMLDoc = New ujXML
                        p 21
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "CO_DOC"
            .chCreate "MessageType"
                .elText = "SALESEORDER"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
'            .elCreateSibling "DestinationAddress"
'                If oPC.TestMode Then
'                    .elText = oPC.EmailAddressForTesting
'                    pDestinationEmailAddress = oPC.EmailAddressForTesting
'                Else
'                    If Me.Supplier Is Nothing Then
'                        .elText = ""
'                        pDestinationEmailAddress = ""
'                    Else
'                        .elText = Me.Supplier.OrderToAddress.Email
'                        pDestinationEmailAddress = Me.Supplier.OrderToAddress.Email
'                    End If
'                End If
                        p 23
            .elCreateSibling "TemplateName"
                If Me.Customer.SalesOrderTemplateName > "" Then
                    .elText = Me.Customer.SalesOrderTemplateName
                Else
                    .elText = "CO_DOC"
                End If
             strTemplateName = .elText
               
            .elCreateSibling "SendersEmail"
                If oPC.EmailFrom > "" Then
                    .elText = oPC.EmailFrom
                Else
                    .elText = Me.SendersEmail
                End If
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
            .elCreateSibling "DocDate", True
                .elText = Me.DOCDate
            .elCreateSibling "OurCompany", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "OurAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerName", True
                .elText = Customer.NameAndCode(70)
            .elCreateSibling "CustomerVATNumber", True
                .elText = Customer.VatNumber
            .elCreateSibling "BankingDetails", True
                .elText = Replace(oPC.Configuration.DefaultCompany.BankDetails, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerWithAddress", True
                If Customer.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Customer.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
            .elCreateSibling "CustomerPhone", True
                If Customer.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.Customer.BillTOAddress.Phone > "", "Phone: " & Customer.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "CustomerFax", True
                If Customer.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Customer.BillTOAddress.Fax > "", "Fax: " & Customer.BillTOAddress.Fax, "")
                End If
            .elCreateSibling "ACNO"
                .elText = IIf(Me.Customer.AcNo > "", "Ac/no. " & Me.Customer.AcNo, "")
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
                            p 24
            
            For i = 1 To Me.COLines.Count
               ' If POLines(i).Fulfilled <> "CAN" Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    .elText = COLines(i).CodeF
                    .elCreateSibling "Title", True
                    If COLines(i).Fulfilled <> "CAN" Then
                       .elText = COLines(i).TitleAuthorPublisher & IIf(COLines(i).ServiceItem, " (" & COLines(i).Qty & " item" & IIf(COLines(i).Qty = 1, ")", "s)"), "")
                    Else
                       .elText = "***CANCELLED***" & COLines(i).TitleAuthorPublisher
                    End If
                    If oPC.AllowsSSInvoicing Then
                        .elCreateSibling "QtyFirm", True
                            .elText = COLines(i).QtyFirmF
                        .elCreateSibling "QtySS", True
                            .elText = COLines(i).QtySSF
                    Else
                        .elCreateSibling "QtyFirm", True
                            .elText = COLines(i).QtyF
                    End If
                    .elCreateSibling "Price", True
                        .elText = COLines(i).PriceF
                    .elCreateSibling "PriceExVat", True
                        .elText = COLines(i).PriceExVatF
                    .elCreateSibling "DiscountRate", True
                        .elText = COLines(i).DiscountF
                    .elCreateSibling "Reference", True
                        .elText = COLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = COLines(i).ExtensionF
                    .elCreateSibling "Note", True
                        .elText = COLines(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = COLines(i).Ext_ExFF
                    .elCreateSibling "VAT", True
                        .elText = COLines(i).VATF
                    .elCreateSibling "VATRate", True
                        .elText = COLines(i).VATRateF
                    .elCreateSibling "DiscountValue", True
                        .elText = COLines(i).DiscF
                    .elCreateSibling "FCName", True
                        If COLines(i).FCID > 0 Then
                            .elText = oPC.Configuration.Currencies.FindCurrencyByID(COLines(i).FCID).Description
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCPrice", True
                        If COLines(i).FCID > 0 Then
                            .elText = COLines(i).ForeignPriceF
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCFactor", True
                        If COLines(i).FCID > 0 Then
                            .elText = COLines(i).FCFactorInvF
                        Else
                            .elText = "n/a"
                        End If
                        
                    .navUP
                    If COLines(i).ExtraCharge > 0 Then
                    .elCreateSibling "DetailLine", True
                        .chCreate "SKU"
                            .elText = ""
                        .elCreateSibling "Title", True
                           .elText = COLines(i).ExtraChargeDescription

                        If oPC.AllowsSSInvoicing Then
                            .elCreateSibling "QtyFirm", True
                                .elText = ""
                            .elCreateSibling "QtySS", True
                                .elText = ""
                        Else
                            .elCreateSibling "QtyFirm", True
                                .elText = ""
                        End If
                        .elCreateSibling "Price", True
                            .elText = COLines(i).ExtraChargeF
                        .elCreateSibling "DiscountRate", True
                            .elText = ""
                        .elCreateSibling "Reference", True
                            .elText = ""
                        .elCreateSibling "Extension", True
                            .elText = COLines(i).ExtraChargeF
                        .elCreateSibling "Note", True
                            .elText = ""
                        .navUP
                    End If
            Next i
                            p 25
            .elCreateSibling "TotalNumberOfLines", True
                .elText = CStr(Me.COLines.Count)
            .elCreateSibling "TotalShipping", True
                .elText = Me.TotalServiceItemF
            .elCreateSibling "TotalDiscount", True
                .elText = Me.TotalDiscountF
            .elCreateSibling "TotalQtyItems", True
                .elText = Me.TotalQtyF
            .elCreateSibling "TotalIncVAT", True
                .elText = Me.TotalPayableF
            .elCreateSibling "TotalExVAT", True
                .elText = Me.TotalNonVATF
            .elCreateSibling "TotalVAT", True
                .elText = Me.TotalVATF
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
                If oPC.Configuration.Staff.FindStaffByID(Me.StaffID) Is Nothing Then
                    .elText = ""
                Else
                    .elText = oPC.Configuration.Staff.FindStaffByID(Me.StaffID).StaffName
                End If
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.OrderText
            .elCreateSibling "VATable", True
                .elText = IIf(Me.VATable, "TRUE", "FALSE")
            .elCreateSibling "ShowVAT", True
                .elText = IIf(Me.ShowVAT, "TRUE", "FALSE")
    End With
                            p 26
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "CO_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With
                            p 27
    If DispatchMode = enMail Then
        'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_HTML.XSLT"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
        
            If fs.FileExists(strWorkingFolder & "CO_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "CO_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "CO_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage 'xMLDoc.docObject.transformNode(objXSL)
            oTF.CloseTextFile
    End If
    
                        p 27
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
        
        strFOFile = strWorkingFolder & "CO_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "CO_" & Me.DOCCode & ".PDF"
                                p 7
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
       ''' strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
'        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
'        F_7_AB_1_ShellAndWaitSimple strCommand
        ChDir "\PBKS\Executables\FOP\"
       'strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)

  LogSaveToFile strCommand

        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
        cnt = 0
        Do While Not fs.FileExists(strPDFFile) And cnt < 40
            MsgWaitObj 1000
            cnt = cnt + 1
        Loop
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
        If fs.FileExists(strPDFFile) Then
       ' MsgBox "here"
            SetFileProperty strPDFFile, "Title", Me.Customer.AcNo
            SetFileProperty strPDFFile, "Author", Me.Customer.Name
        End If
    End If
                        p 8
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    pFilename = strPDFFile

    If DispatchMode <> enMail And DispatchMode <> enView Then
        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    End If
    Set oSM = Nothing
    Set oFSO = Nothing
                        p 9
    oTF.CloseTextFile
    Set oTF = Nothing
    ExportToXML = True
    
    '
'    If DispatchMode = enView Then
'        strExecutable = GetPDFExecutable(strPDFFile)
'        Shell strExecutable & " " & strPDFFile
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
    ErrorIn "a_CO.ExportToXML(DispatchMode,pDestinationEmailAddress,pWholeMessage,pQtyCopies)", _
         Array(DispatchMode, pDestinationEmailAddress, pWholeMessage, pQtyCopies), , , "line number", Array(Erl())
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
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCode
    oTF.WriteToTextFile "DocDate" & vbTab & DOCDate
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.DefaultCompany.CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "CustomerName" & vbTab & Customer.NameAndCode(70)
    oTF.WriteToTextFile "ACNO" & vbTab & IIf(Me.Customer.AcNo > "", "Ac/no. " & Me.Customer.AcNo, "")
    oTF.WriteToTextFile "CustomerVATNumber" & vbTab & Customer.VatNumber
    oTF.WriteToTextFile "BankingDetails" & vbTab & Replace(oPC.Configuration.DefaultCompany.BankDetails, Chr(13) & Chr(10), Chr(10))
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

    s = "SKU" & vbTab & "Title" & vbTab & "QtyFirm" & vbTab & "QtySS" & vbTab & "Price" & vbTab & "PriceExVat" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab _
    & "Extension" & vbTab & "Note" & vbTab & "ExtensionExVat" & vbTab & "ExtensionVATOnly" & vbTab & "VATRate" & vbTab & "DiscountValue" & vbTab _
    & "FCName" & vbTab & "FCPrice" & vbTab & "FCFactor"
    
    oTF.WriteToTextFile s
               
                
  lngNumberOfLines = 0
  For i = 1 To COLines.Count
    lngNumberOfLines = lngNumberOfLines + 1
    s = COLines(i).CodeF & vbTab
        If COLines(i).Fulfilled <> "CAN" Then
            s2 = COLines(i).TitleAuthorPublisher & IIf(COLines(i).ServiceItem, " (" & COLines(i).Qty & " item" & IIf(COLines(i).Qty = 1, ")", "s)"), "")
        Else
            s2 = "***CANCELLED***" & COLines(i).TitleAuthorPublisher
        End If
    s = s & s2
    s = s & vbTab & COLines(i).QtyFirmF
    s = s & vbTab & COLines(i).QtySSF
    s = s & vbTab & COLines(i).PriceF
    s = s & vbTab & COLines(i).PriceExVatF
    s = s & vbTab & COLines(i).DiscountF
    s = s & vbTab & COLines(i).Ref
    s = s & vbTab & COLines(i).ExtensionF
    s = s & vbTab & COLines(i).Note
    s = s & vbTab & COLines(i).Ext_ExFF
    s = s & vbTab & COLines(i).VATF
    s = s & vbTab & COLines(i).VATRateF
    s = s & vbTab & COLines(i).DiscF
                        If COLines(i).FCID > 0 Then
                            s2 = oPC.Configuration.Currencies.FindCurrencyByID(COLines(i).FCID).Description
                        Else
                            s2 = "n/a"
                        End If

                s = s & vbTab & s2
                        If COLines(i).FCID > 0 Then
                            s2 = COLines(i).ForeignPriceF
                        Else
                            s2 = "n/a"
                        End If
                s = s & vbTab & s2
                        If COLines(i).FCID > 0 Then
                            s2 = COLines(i).FCFactorInvF
                        Else
                            s2 = "n/a"
                        End If
                
                s = s & vbTab & s2
    oTF.WriteToTextFile s
Next

    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(Me.COLines.Count)
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.WriteToTextFile "StaffMember" & vbTab & IIf(Me.StaffName > "", Me.StaffName, "")
    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalPayableF
    oTF.WriteToTextFile "TotalExVAT" & vbTab & TotalNonVATF
    oTF.WriteToTextFile "TotalVAT" & vbTab & TotalVATF
    oTF.WriteToTextFile "VATable" & vbTab & IIf(Me.VATable, "TRUE", "FALSE")
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)

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
    ErrorIn "a_PO.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub


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
    ErrorIn "a_CO.TranslateErrors(pRawors)", pRawors
End Function
Public Property Get Log() As String
    On Error GoTo errHandler
    Log = Trim(mudtProps.Log)
    Exit Property
errHandler:
    ErrorIn "a_CO.Log"
End Property

Public Property Let Log(val As String)
    On Error GoTo errHandler
    mudtProps.Log = Trim(val)
    Exit Property
errHandler:
    ErrorIn "a_CO.Log(val)", val
End Property

Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.LoadClassorsCollection"
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

    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function
Public Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetDirty(pVal)", pVal
End Sub


Public Function Post() As String
    On Error GoTo errHandler
Dim strResult As String
Dim oSM As z_StockManager
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    'oPC.COShort.BeginTrans
    Me.ApplyEdit strResult
    
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
            Case stCOMPLETE, stISSUED
                oSM.IssueCO Me
              '  SetStatus stCOMPLETE
            Case stPROFORMA
                oSM.IssueCO Me
        '        SetStatus stPROFORMA
            End Select
    Case stISSUED
            Select Case Status
            Case stCANCELLED
                oSM.CancelCO Me
            Case stISSUED, stCOMPLETE
                oSM.IssueCO Me
            End Select
            Post = "ISSUED"
            
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    Case stPROFORMA      'Issued PROFORMA
            Select Case Status
            Case stVOID
                oSM.CancelCO Me
        '        SetStatus stVOID
            Case stInProcess
                Post = "OR: was proforma, can't now be inprocess."
            Case stCOMPLETE
                oSM.IssueCO Me
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    Set oSM = Nothing
   ' oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Post", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetStatus(val)", val
End Sub
Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit strResult
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.VoidDocument"
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
    ErrorIn "a_CO.Reload"
End Sub
Public Property Get CanCancel() As Boolean
    On Error GoTo errHandler
Dim bCanCancel As Boolean
Dim oCOL As a_COL
    bCanCancel = True
    For Each oCOL In COLines
        If oCOL.Fulfilled <> "CAN" Then
            If oCOL.QtyDispatched > 0 Then
                bCanCancel = False
                Exit For
            End If
        End If
    Next
    CanCancel = bCanCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.CanCancel"
End Property

Public Sub PasteLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, VATRate As Double, _
        Ref As String, ExtraChargePID As String, ExtraChargeValue As Long, ETA As Date, _
        FCPrice As Long, FCFactor As Double, FCID As Long)
Dim oSM As New z_StockManager

    oSM.PasteCOLine PID, QtyFirm, QtySS, Price, DiscountRate, VATRate, mudtProps.TRID, Ref, ExtraChargePID, ExtraChargeValue, ETA, FCPrice, FCFactor, FCID
    
End Sub

Public Sub setDelToAddressImmediate(val As a_Address)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    Set oDelToAddress = val
    mudtProps.GoodsAddressID = oDelToAddress.ID
    oSM.SetAddressImmediate "DEL", Me.TRID, oDelToAddress.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.setDelToAddressImmediate(val)", val
End Sub
Public Sub SetBillToAddressImmediate(val As a_Address)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    Set oBillToAddress = val
    mudtProps.BillToAddressID = oBillToAddress.ID
    oSM.SetAddressImmediate "BILL", Me.TRID, oBillToAddress.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SetBillToAddressImmediate(val)", val
End Sub

'Public Property Get SendersEmail()
'    SendersEmail = FNS(mudtProps.StaffEmail)
'End Property
Public Property Get SendersEmail() As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
        SendersEmail = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).EMail
    If strT > "" Then
        SendersEmail = strT
    Else
        SendersEmail = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.SendersEmail"
End Property

Public Property Get TotalServiceItem() As String
    On Error GoTo errHandler

        TotalServiceItem = mudtProps.TotalServiceItem ' / oPC.Configuration.DefaultCurrency.Divisor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalServiceItem(bForeign)"
End Property
Public Property Get TotalServiceItemF() As String
    On Error GoTo errHandler
        TotalServiceItemF = Format(TotalServiceItem / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalServiceItemF(bForeign)"
End Property
Public Property Get TotalDiscount() As Long
    On Error GoTo errHandler
        TotalDiscount = mudtProps.TotalDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalDiscount(bForeign)"
End Property
Public Property Get TotalDiscountF() As String
    On Error GoTo errHandler
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalDiscountF(bForeign)"
End Property
Public Property Get TotalNonVATF() As String
    On Error GoTo errHandler
        TotalNonVATF = Format(mudtProps.TotalNonVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.TotalNonVATF(bForeign)"
End Property

