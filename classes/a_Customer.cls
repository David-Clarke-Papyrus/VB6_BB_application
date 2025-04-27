VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(strMsg As String)
Event DefaultAddressChanged()
Event Dirty(bYesNO As Boolean)
Event PossibleDuplicates(Duplicates As c_Customer)
Private mudtProps As CustomerProps
Private mcolStack As Collection
Private WithEvents mcolAddresses As ch_Addresses
Attribute mcolAddresses.VB_VarHelpID = -1
Private mcolIG As ch_IG
Attribute mcolIG.VB_VarHelpID = -1
Private mcolCC As ch_IG
Private oApproAddress As a_Address
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oOrderToAddress As a_Address
Private xMLDoc As ujXML

Private tlTitles As z_TextList
Private tlInterestGroupsActive As z_TextList
Private tlInterestGroupsAll As z_TextList
Private tlInterestGroupsLoyalty As z_TextList
Private tlCustomerTypesActive As z_TextList
Private tlCustomerTypesAll As z_TextList
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mApproAddressIdx As String
Private mBilltoAddressIdx As String
Private mDelToAddressIdx As String
Private mOrderToAddressIdx As String
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim mControl As String
Dim bReadyForDuplicateCheck As Boolean
Public Property Let ParentCustomerID(val As Long)
    mudtProps.ParentCustomerID = val
End Property
Public Property Get ParentCustomerID() As Long
    ParentCustomerID = mudtProps.ParentCustomerID
End Property
Public Property Get ParentCustomerName() As String
    ParentCustomerName = mudtProps.ParentCustomerName
End Property

Public Property Get InterestGroups() As ch_IG
    Set InterestGroups = mcolIG
End Property
Public Property Get CustomerTypes() As ch_IG
    Set CustomerTypes = mcolCC
End Property
Public Property Get IsBookClub() As Boolean
Dim oCT As a_IG
    
    IsBookClub = False
    For Each oCT In CustomerTypes
        If oCT.IGID = oPC.Configuration.IGBookClubID Then
            IsBookClub = True
            Exit For
        End If
    Next
End Property
Private Sub mcolAddresses_FirstAddressAdded()
    On Error GoTo errHandler
    SetBillToAddressidx Me.Addresses(1).Key
    SetApproAddressidx Me.Addresses(1).Key
    SetDelToAddressidx Me.Addresses(1).Key
    SetOrderToAddressidx Me.Addresses(1).Key
    RaiseEvent DefaultAddressChanged
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mcolAddresses_FirstAddressAdded"
End Sub
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As CustomerData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
Dim udtData As CustomerData
    udtData.buffer = buffer
    LSet mudtProps = udtData
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetState(buffer)", buffer
End Sub
'Public Property Get IGs() As c_IGs
'    Set IGs = cIGs
'End Property
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "Addresses", mcolAddresses.GetSuperState
    .WriteProperty "IGs", mcolIG.GetSuperState
    .WriteProperty "CCs", mcolCC.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GetSuperState"
End Function

Public Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim oAdd As a_Address
Dim i As Integer

    If buffer = "" Then Exit Sub
    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      SetState .ReadProperty("State")
      mcolIG.SetSuperState .ReadProperty("IGs")
      mcolCC.SetSuperState .ReadProperty("CCs")
      mcolAddresses.SetSuperState .ReadProperty("Addresses")
    End With
    
    For i = 1 To Addresses.Count
        If Addresses(i).Appro Then
            mApproAddressIdx = i
            Set oApproAddress = Nothing
            Set oApproAddress = Addresses(i)
        End If
        If Addresses(i).BillTo Then
            mBilltoAddressIdx = i
            Set oBillToAddress = Nothing
            Set oBillToAddress = Addresses(i)
        End If
        If Addresses(i).DelTo Then
            mDelToAddressIdx = i
            Set oDelToAddress = Nothing
            Set oDelToAddress = Addresses(i)
        End If
        If Addresses(i).OrderTo Then
            mOrderToAddressIdx = i
            Set oOrderToAddress = Nothing
            Set oOrderToAddress = Addresses(i)
        End If
    Next i
    If oApproAddress Is Nothing And Addresses.Count > 0 Then
        Set oApproAddress = Nothing
        Set oApproAddress = Addresses(1)
        oApproAddress.BeginEdit
        oApproAddress.Appro = True
        oApproAddress.ApplyEdit
    End If
    If oBillToAddress Is Nothing And Addresses.Count > 0 Then
        Set oBillToAddress = Nothing
        Set oBillToAddress = Addresses(1)
        oBillToAddress.BeginEdit
        oBillToAddress.BillTo = True
        oBillToAddress.ApplyEdit
    End If
    If oDelToAddress Is Nothing And Addresses.Count > 0 Then
        Set oDelToAddress = Nothing
        Set oDelToAddress = Addresses(1)
        oDelToAddress.BeginEdit
        oDelToAddress.DelTo = True
        oDelToAddress.ApplyEdit
    End If
    If oOrderToAddress Is Nothing And Addresses.Count > 0 Then
        Set oOrderToAddress = Nothing
        Set oOrderToAddress = Addresses(1)
        oOrderToAddress.BeginEdit
        oOrderToAddress.OrderTo = True
        oOrderToAddress.ApplyEdit
    End If

  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetSuperState(buffer)", buffer
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolAddresses.BeginEdit
  mcolIG.BeginEdit
  mcolCC.BeginEdit
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
    mcolAddresses.CancelEdit
    mcolIG.CancelEdit
    mcolCC.CancelEdit
    mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CancelEdit"
End Sub

Public Sub ApplyEdit(lngResult As Long)
    On Error GoTo errHandler
Dim strTemp As String
Dim OpenResult As Integer
Dim strPos As String

  Dim objPersist As a_Customer_P
    lngResult = 0
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    If mcolStack.Count = 0 Then Err.Raise 445
    If Not oBillToAddress Is Nothing Then
      If oBillToAddress.Phone > "" Then
          Me.SetPhone oBillToAddress.Phone
      End If
    End If
    mcolAddresses.ApplyEdit
    mcolIG.ApplyEdit
    mcolCC.ApplyEdit

    Set objPersist = New a_Customer_P
    If mudtProps.IsDeleted Then
      objPersist.DeleteObject mudtProps.CustID
      mcolStack.Remove mcolStack.Count
      mudtProps.IsNew = True
      mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
      If Not IsValid Then Err.Raise 445
      strTemp = objPersist.Save(GetSuperState)
      If strTemp = "DUPLICATE" Then
          oPC.COShort.RollbackTrans
          mcolAddresses.BeginEdit
          mcolIG.BeginEdit
          mcolCC.BeginEdit
          lngResult = 22
          GoTo EXIT_Handler
      ElseIf strTemp > "" Then
          
          mcolStack.Remove mcolStack.Count
          SetSuperState strTemp
          mudtProps.IsNew = False
      End If
    Else
      mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    mudtProps.IsDirty = False
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Sub

errHandler:
    ErrPreserve
    lngResult = 101
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ApplyEdit(lngResult)", lngResult, , , "Line number,BrokenRules", Array(Erl(), mobjValid.AllBrokenRules)
End Sub
Public Function CustomerIndexClashes() As Boolean
Dim oSM As New a_Customer_P
    
    If Me.IsNew Then
        CustomerIndexClashes = oSM.CustomerIndexClashes(0, Me.AcNo)
    Else
        CustomerIndexClashes = oSM.CustomerIndexClashes(Me.ID, Me.AcNo)
    End If

End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mcolIG = New ch_IG
    Set mcolCC = New ch_IG
    Set mcolAddresses = New ch_Addresses
    mcolAddresses.SetCustomerParent Me
    Set tlTitles = New z_TextList
    Set tlInterestGroupsActive = New z_TextList
    Set tlInterestGroupsAll = New z_TextList
    Set tlCustomerTypesActive = New z_TextList
    Set tlCustomerTypesAll = New z_TextList
    Set tlInterestGroupsLoyalty = New z_TextList
'    Set cIGs = New c_IGs
    tlInterestGroupsLoyalty.Load ltInterestGroupLoyalty
    tlInterestGroupsActive.Load ltInterestGroupActive
    tlInterestGroupsAll.Load ltInterestGroupAll
    tlCustomerTypesActive.Load ltCustomerTypeActive
    tlCustomerTypesAll.Load ltCustomerTypeAllExceptLoyalty
    mudtProps.CustomerTypeID = oPC.Configuration.DefaultCT
    mudtProps.IsNew = True
    Set colClassors = New Collection
    LoadClassorsCollection
    mBilltoAddressIdx = 1
  '  mobjValid.RuleBroken "MINIMUM", True
    mobjValid.RuleBroken "NAME", True
    VATable = True
    bReadyForDuplicateCheck = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Class_Initialize"
End Sub
Public Sub InitializeNewCustomer(pType As enumCustomerType)
Dim oAdd As a_Address
    If pType = enBusiness Then
        mudtProps.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID
    ElseIf pType = enPrivate Then
        mudtProps.CustomerTypeID = oPC.Configuration.PrivateCustomerTypeID
    ElseIf pType = enBookclub Then
        mudtProps.CustomerTypeID = oPC.Configuration.BookClubCustomerTypeID
    End If
'    Me.CustomerTypes.Add
    Set oAdd = Addresses.Add
    oAdd.BeginEdit
    oAdd.SetDescription "Default"
    oAdd.ApplyEdit
End Sub
Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set mcolAddresses = Nothing
    Set mcolIG = Nothing
    Set mcolCC = Nothing
    Set tlTitles = Nothing
  '  Set cIGs = Nothing
    Set tlInterestGroupsActive = Nothing
    Set tlInterestGroupsLoyalty = Nothing
    Set tlCustomerTypesActive = Nothing
    Set tlInterestGroupsAll = Nothing
    Set tlCustomerTypesAll = Nothing
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Class_Terminate"
End Sub
Public Property Get CustomerCategory() As Integer
    CustomerCategory = FNN(mudtProps.CustomerCategory)
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mobjValid_RuleUnbroken(ors)", oRS
End Sub

'Private Sub mobjValid_ThisRuleUnbroken(Thisor As String)
'    On Error GoTo ErrHandler
'    If Thisor = "MINIMUM" Then 'we can check for duplicates
'        LookforDuplicates
'    End If
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.mobjValid_ThisRuleUnbroken(ors)", ors
'End Sub
Public Function LookforDuplicates()
Dim oDup As New c_Customer
Dim bFound As Boolean
Dim tmpCust As a_Customer
Dim oAdd As a_Address

    bFound = False
    For Each oAdd In Me.Addresses
        If oAdd.Phone > "" Then
            oDup.LoadDuplicates bFound, StripToNumerics(PhoneFormat(oAdd.Phone, oPC.DefaultAreaCode)), oAdd.EMail, StripToNumerics(Me.MOBILE), mudtProps.CustID
        End If
    Next
    If Me.MOBILE > "" Then
            oDup.LoadDuplicates bFound, "", "", StripToNumerics(Me.MOBILE), mudtProps.CustID
    End If
    If bFound Then
        RaiseEvent PossibleDuplicates(oDup)
    End If
        
End Function

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mobjValid_Status(pMsg)", pMsg
End Sub
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
    ErrorIn "a_Customer.GetSTatus"
End Function
Public Property Get InterestGroupsLoyalty_tl() As z_TextList
    Set InterestGroupsLoyalty_tl = tlInterestGroupsLoyalty
End Property

Public Property Get InterestGroupsAll_tl() As z_TextList
    Set InterestGroupsAll_tl = tlInterestGroupsAll
End Property
Public Property Get InterestGroupsActive_tl() As z_TextList
    Set InterestGroupsActive_tl = tlInterestGroupsActive
End Property
Public Property Get CustomerTypesALL_tl() As z_TextList
    Set CustomerTypesALL_tl = tlCustomerTypesAll
End Property
Public Property Get CustomerTypesActive_tl() As z_TextList
    Set CustomerTypesActive_tl = tlCustomerTypesActive
End Property
Public Function Reload()
Dim lngID As Long
    lngID = Me.ID
    Class_Terminate
    Class_Initialize
    Load lngID, True

End Function
Public Function Load(Optional CustomerID As Long = 0, Optional pACNO As String = "", Optional pDefaultPhone As String = "", Optional pSAN As String) As Boolean
    On Error GoTo errHandler
Dim objPersist As a_Customer_P
Dim tmpBuffer As String
Dim i As Integer
    If mcolStack.Count > 0 Then Err.Raise 445
 '   If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    mBilltoAddressIdx = 1 'just in case
    
    ' code to load the object goes here
    Set objPersist = New a_Customer_P
    Set oBillToAddress = Nothing
    Load = True
    tmpBuffer = ""
    If CustomerID > 0 Then
        tmpBuffer = objPersist.Fetch(CustomerID)
        If tmpBuffer > "" Then
            SetSuperState tmpBuffer
        End If
    ElseIf pACNO > "" Then
        tmpBuffer = objPersist.Fetch(, pACNO)
        If tmpBuffer > "" Then
            SetSuperState tmpBuffer
        End If
    ElseIf pSAN > "" Then
        tmpBuffer = objPersist.Fetch(, , , pSAN)
        If tmpBuffer > "" Then
            SetSuperState tmpBuffer
        End If
    ElseIf pDefaultPhone > "" Then
        tmpBuffer = objPersist.Fetch(, , pDefaultPhone)
        If tmpBuffer > "" Then
            SetSuperState tmpBuffer
        End If
    End If
    If tmpBuffer = "" Then
        Load = False
        Exit Function
    Else
'        cIGs.Load mudtProps.CustID
        Set objPersist = Nothing
        mobjValid.RuleBroken "Name", False
        mobjValid.RuleBroken "Phone", False
        mobjValid.RuleBroken "TP", False
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Load(CustomerID,pACNO,pDefaultPhone)", Array(CustomerID, pACNO, _
         pDefaultPhone)
End Function

Public Property Get CustomerID() As Long
    On Error GoTo errHandler
    CustomerID = mudtProps.CustID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustomerID"
End Property

Public Sub SetApproAddressidx(val As String)
Dim oAdd As a_Address
    On Error GoTo errHandler
    mApproAddressIdx = val
    Set oApproAddress = Nothing
    Set oApproAddress = Me.Addresses(mApproAddressIdx)
    For Each oAdd In Addresses
        oAdd.Appro = False
    Next
    Addresses(mApproAddressIdx).Appro = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetApproAddressidx(val)", val
End Sub
Public Property Get ApproAddressIdx() As Integer
    ApproAddressIdx = mApproAddressIdx
End Property
Public Property Get ApproAddress() As a_Address
    Set ApproAddress = oApproAddress
End Property

Public Sub SetBillToAddressidx(val As String)
Dim oAdd As a_Address

    On Error GoTo errHandler
    mBilltoAddressIdx = val
    Set oBillToAddress = Nothing
    Set oBillToAddress = Addresses(mBilltoAddressIdx)
    For Each oAdd In Addresses
        oAdd.BillTo = False
    Next
    Addresses(mBilltoAddressIdx).BillTo = True

    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetBillToAddressidx(val)", val
End Sub
Public Property Get BillToAddressIdx() As Integer
    BillToAddressIdx = mBilltoAddressIdx
End Property
Public Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property

Public Sub SetDelToAddressidx(val As String)
Dim oAdd As a_Address
    On Error GoTo errHandler
    mDelToAddressIdx = val
    Set oDelToAddress = Nothing
    Set oDelToAddress = Me.Addresses(mDelToAddressIdx)
    For Each oAdd In Addresses
        oAdd.DelTo = False
    Next
    Addresses(mDelToAddressIdx).DelTo = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetDelToAddressidx(val)", val
End Sub
Public Property Get DelTOAddressIdx() As Integer
    DelTOAddressIdx = mDelToAddressIdx
End Property
Public Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property

Public Sub SetOrderToAddressidx(val As String)
Dim oAdd As a_Address
    On Error GoTo errHandler
    mOrderToAddressIdx = val
    Set oOrderToAddress = Nothing
    Set oOrderToAddress = Me.Addresses(mOrderToAddressIdx)
    For Each oAdd In Addresses
        oAdd.OrderTo = False
    Next
    Addresses(mOrderToAddressIdx).OrderTo = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetOrderToAddressidx(val)", val
End Sub
Public Property Get OrderToAddressIdx() As Integer
    OrderToAddressIdx = mOrderToAddressIdx
End Property
Public Property Get OrderToAddress() As a_Address
    Set OrderToAddress = oOrderToAddress
End Property


Public Sub DeleteCustomer()
    On Error GoTo errHandler
    mudtProps.IsDeleted = True
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.DeleteCustomer"
End Sub
Public Property Let CustomerTypeID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CustomerTypeID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustomerTypeID(val)", val
End Property
Public Property Get CustomerTypeID() As Long
    On Error GoTo errHandler
    CustomerTypeID = mudtProps.CustomerTypeID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustomerTypeID"
End Property
Public Property Get CustomerTypeShortname() As String
    On Error GoTo errHandler
    CustomerTypeShortname = tlCustomerTypesAll.Item(mudtProps.CustomerTypeID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustomerTypeShortname"
End Property
Public Property Let GetsCatalogue(val As Boolean)
    On Error GoTo errHandler
    mudtProps.GetsCatalogue = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GetsCatalogue(val)", val
End Property
Public Property Get GetsCatalogue() As Boolean
    On Error GoTo errHandler
    GetsCatalogue = mudtProps.GetsCatalogue
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GetsCatalogue"
End Property
Public Property Let OneLinePerInvoice(val As Boolean)
    On Error GoTo errHandler
    mudtProps.OneLinePerInvoice = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.OneLinePerInvoice(val)", val
End Property
Public Property Get OneLinePerInvoice() As Boolean
    On Error GoTo errHandler
    OneLinePerInvoice = mudtProps.OneLinePerInvoice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.OneLinePerInvoice"
End Property
Public Property Get CompleteOrder() As Boolean
    On Error GoTo errHandler
    CompleteOrder = mudtProps.CompleteOrder
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CompleteOrder"
End Property

Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.VATable(val)", val
End Property
Public Property Get VATable() As Boolean
    On Error GoTo errHandler
    VATable = mudtProps.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.VATable"
End Property
Public Property Let ShowVAT(val As Boolean)
    On Error GoTo errHandler
    mudtProps.ShowVAT = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ShowVAT(val)", val
End Property
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ShowVAT"
End Property
Public Function SetRepID(val As Long)
    
    mudtProps.SalesRepID = val
    mudtProps.IsDirty = True

End Function
Public Function repid() As Long
    repid = FNN(mudtProps.SalesRepID)
End Function
Public Property Get Repname() As String
    Repname = FNS(mudtProps.Repname)
End Property
Public Property Let Repname(val As String)
    mudtProps.Repname = val
End Property
Public Property Get GenerateSeparateInvoicesForSeparateOrders() As Boolean
    GenerateSeparateInvoicesForSeparateOrders = FNB(mudtProps.GenerateSeparateInvoicesForSeparateOrders)
End Property
Public Property Let GenerateSeparateInvoicesForSeparateOrders(val As Boolean)
    mudtProps.GenerateSeparateInvoicesForSeparateOrders = val
End Property
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.CustID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ID"
End Property
Public Property Get DefaultDiscount() As Double
    On Error GoTo errHandler
    DefaultDiscount = mudtProps.DefaultDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.DefaultDiscount"
End Property
Public Property Get DefaultDiscountF() As String
    On Error GoTo errHandler
    DefaultDiscountF = PBKSPercentF(mudtProps.DefaultDiscount)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.DefaultDiscountF"
End Property
Public Function SetDefaultDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.DefaultDiscount
    bOK = SetField_DOUBLE(mudtProps.DefaultDiscount, val, "DISCOUNT", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.DefaultDiscount) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject(val, "DISCOUNT")
    End If
    SetDefaultDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetDefaultDiscount(val)", val
End Function
'Public Property Get Titles() As z_TextList
'    Set Titles = tlTitles
'End Property
Public Property Get CreditLimit() As Long
    On Error GoTo errHandler
    CreditLimit = FNN(mudtProps.CreditLimit)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetCreditLimit"
End Property
Public Property Get CreditLimitF() As String
    On Error GoTo errHandler
    CreditLimitF = Format(mudtProps.CreditLimit, "" & oPC.Configuration.DefaultCurrency.Symbol & "#,##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetCreditLimitF"
End Property
Public Function SetCreditLimit(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim Temp As Long
    Temp = mudtProps.CreditLimit
    bOK = SetField_LONG(mudtProps.CreditLimit, val, "CREDITLIMIT", mcolStack.Count)
    If bOK And (Temp <> mudtProps.CreditLimit) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject(val, "CREDITLIMIT")
    End If
    SetCreditLimit = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetCreditLimit(val)", val
End Function
Public Sub SetPaymentStyle(val As String)
    mudtProps.PaymentStyle = val
End Sub
Public Property Get PaymentStyle() As String
    PaymentStyle = FNS(mudtProps.PaymentStyle)
End Property

Public Sub SetTerms(val As Integer)
    mudtProps.Terms = val
End Sub
Public Property Get Terms() As Integer
    Terms = FNN(mudtProps.Terms)
End Property
Public Property Get TermsF() As String
    Select Case mudtProps.Terms
    Case 0
        TermsF = "Current"
    Case 15
        TermsF = "15 days"
    Case 30
        TermsF = "30 days"
    Case 60
        TermsF = "60 days"
    Case 90
        TermsF = "90 days"
    Case 120
        TermsF = "120 days"
    End Select
End Property
Public Sub SetBalances(cur As Double, Thirty As Double, Sixty As Double, Ninety As Double, OneTwenty As Double, Total As Double)
    mudtProps.Bal120 = OneTwenty
    mudtProps.Bal30 = Thirty
    mudtProps.Bal60 = Sixty
    mudtProps.Bal90 = Ninety
    mudtProps.BalCur = cur
    mudtProps.Balance = Total
End Sub
Public Property Get oLineToInvoice() As Boolean
    oLineToInvoice = FNB(mudtProps.oLineToInvoice)
End Property
Public Property Get Blocked() As Boolean
    Blocked = FNB(mudtProps.Blocked)
End Property
Public Property Get UseQuotedPrice() As Boolean
    UseQuotedPrice = FNB(mudtProps.UseQuotedPrice)
End Property
Public Property Get Balance() As Double
    Balance = FNDBL(mudtProps.Balance)
End Property
Public Property Get BalanceF() As String
    BalanceF = Format(mudtProps.Balance, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BalanceCur() As Double
    BalanceCur = FNDBL(mudtProps.BalCur)
End Property
Public Property Get BalanceCurF() As String
    BalanceCurF = Format(mudtProps.BalCur, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Balance30() As Double
    Balance30 = FNDBL(mudtProps.Bal30)
End Property
Public Property Get Balance30F() As String
    Balance30F = Format(mudtProps.Bal30, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Balance60() As Double
    Balance60 = FNDBL(mudtProps.Bal60)
End Property
Public Property Get Balance60F() As String
    Balance60F = Format(mudtProps.Bal60, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Balance90() As Double
    Balance90 = FNDBL(mudtProps.Bal90)
End Property
Public Property Get Balance90F() As String
    Balance90F = Format(mudtProps.Bal90, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Balance120() As Double
    Balance120 = FNDBL(mudtProps.Bal120)
End Property
Public Property Get Balance120Plus() As Double
    Balance120Plus = FNDBL(mudtProps.Bal120Plus)
End Property
Public Property Get Balance120F() As String
    Balance120F = Format(mudtProps.Bal120, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Balance120PlusF() As String
    Balance120PlusF = Format(mudtProps.Bal120Plus, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BalBF() As Double
    BalBF = Format(Me.Balance120Plus + Me.Balance120 + Me.Balance90 + Me.Balance60 + Me.Balance30, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BalBFF() As String
    BalBFF = Format(Me.BalBF, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Titles() As z_TextList
    Set Titles = tlTitles
End Property
Public Function OKForDeletion(pXASafe As XArrayDB, pXAUnSafe As XArrayDB, pDPTP As c_DocsPerTP) As Boolean
    On Error GoTo errHandler
Dim oACt As New c_DocsPerTP
Dim i As Long

    OKForDeletion = True
    Set XA = Nothing
    Set XA = pXASafe
    oACt.Load Me.ID, False
    XA.ReDim 1, oACt.Count, 1, 7
    If oACt.Count > 0 Then
        OKForDeletion = False
        For i = 1 To oACt.Count
            XA(i, 1) = oACt.Item(i).DocDateF
            XA(i, 2) = oACt.Item(i).DOCCode
            XA(i, 3) = oACt.Item(i).DocTypeF
            XA(i, 4) = oACt.Item(i).DocStatus
          '  XA(i, 5) = oACt.Item(i).DocCode
            XA(i, 6) = oACt.Item(i).DOCID
            XA(i, 7) = oACt.Item(i).DocDateForSort
        Next
        Set pDPTP = oACt
    End If
    
    Set XB = Nothing
    Set XB = pXAUnSafe
    oACt.Load Me.ID, True
    XB.ReDim 1, oACt.Count, 1, 7
    If oACt.Count > 0 Then
        OKForDeletion = False
        For i = 1 To oACt.Count
            XB(i, 1) = oACt.Item(i).DocDateF
            XB(i, 2) = oACt.Item(i).DOCCode
            XB(i, 3) = oACt.Item(i).DocType
            XB(i, 4) = oACt.Item(i).DocStatus
          '  XA(i, 5) = oACt.Item(i).DocCode
            XB(i, 6) = oACt.Item(i).DOCID
            XB(i, 7) = oACt.Item(i).DocDateForSort
        Next
        Set pDPTP = oACt
        OKForDeletion = False

    End If
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.OKForDeletion(pXASafe,pXAUnSafe,pDPTP)", Array(pXASafe, pXAUnSafe, pDPTP)
End Function


Public Property Get CustNotifyBookLaunch() As Boolean
    On Error GoTo errHandler
  CustNotifyBookLaunch = FNB(mudtProps.CustNotifyBookLaunch)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustNotifyBookLaunch"
End Property
Public Property Let CustNotifyBookLaunch(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CustNotifyBookLaunch = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustNotifyBookLaunch(val)", val
End Property
Public Property Get CustNotifyBookPromotion() As Boolean
    On Error GoTo errHandler
  CustNotifyBookPromotion = FNB(mudtProps.CustNotifyBookPromotion)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustNotifyBookPromotion"
End Property
Public Property Let CustNotifyBookPromotion(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CustNotifyBookPromotion = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustNotifyBookPromotion(val)", val
End Property

Public Property Get CustNotifyBookSale() As Boolean
    On Error GoTo errHandler
  CustNotifyBookSale = FNB(mudtProps.CustNotifyBookSale)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustNotifyBookSale"
End Property
Public Property Let CustNotifyBookSale(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CustNotifyBookSale = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustNotifyBookSale(val)", val
End Property




Public Property Get CanBeDeleted() As Boolean
    On Error GoTo errHandler
  CanBeDeleted = FNB(mudtProps.CanBeDeleted)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CanBeDeleted"
End Property
Public Property Let CanBeDeleted(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CanBeDeleted = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CanBeDeleted(val)", val
End Property

Public Property Get Addresses() As ch_Addresses
    Set Addresses = mcolAddresses
End Property
Public Property Get Initials() As String
    On Error GoTo errHandler
  Initials = FNS(mudtProps.Initials)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Initials"
End Property
Public Function SetInitials(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetInitials = True
    If Len(val) > Len(mudtProps.Initials) Then
        Err.Raise 384
    End If
    mudtProps.Initials = val
    SetInitials = ValidateObject(val, "INITIALS")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get Phone() As String
    On Error GoTo errHandler
    If Not BillTOAddress Is Nothing Then
        Phone = BillTOAddress.Phone
    Else
        Phone = FNS(mudtProps.Phone)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Phone"
End Property
Public Property Get Phonef() As String
    If Len(Phone) = 10 Then
        Phonef = Format(Phone, "\(@@@\)\ @@@\-@@@@")
    Else
        Phonef = Phone
    End If
End Property

Public Property Get DispatchMethod() As String
    DispatchMethod = FNS(mudtProps.DispatchMethod)
End Property
Public Function SetDispatchMethod(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    bOK = IIf(UCase(val) = "E" Or UCase(val) = "M" Or UCase(val) = "P", True, False)
    If bOK Then
        mudtProps.DispatchMethod = val
    End If
    SetDispatchMethod = bOK
End Function

Public Property Get SearchPhone() As String
    On Error GoTo errHandler
    SearchPhone = StripToNumerics(FNS(mudtProps.SearchPhone))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SearchPhone"
End Property
Private Function CleanForSearchField(pIn As String) As String
Dim str As String
Dim i As Integer
    i = InStr(1, pIn, "@")
    If i > 0 Then 'we have an email
        str = Left(pIn, i - 1)
    Else  'we have a phone
        str = Right(StripToNumerics(pIn), 6)
    End If
    CleanForSearchField = str
End Function
Public Function SetPhone(val As String) As Boolean
    On Error GoTo errHandler
Dim strSearchPhone As String
'This is used to store either phone or email
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetPhone = True
    If Len(val) > Len(mudtProps.Phone) Then
        Err.Raise 384
    End If
    mudtProps.Phone = val
    Me.BillTOAddress.SetPhone val
    bOK = ValidateObject(val, "PHONE")
    mudtProps.SearchPhone = CleanForSearchField(val)  'Right(StripToNumerics(val), 6)
    If bOK Then
        mudtProps.IsDirty = True
    End If
    SetPhone = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetPhone(val)", val
End Function
Public Function SetContactpersonPhone(val As String) As Boolean
    On Error GoTo errHandler
Dim strSearchPhone As String
'This is used to store either phone or email
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetContactpersonPhone = True
    If Len(val) > Len(mudtProps.ContactpersonPhone) Then
        Err.Raise 384
    End If
    mudtProps.ContactpersonPhone = val
    bOK = ValidateObject(val, "")
    If bOK Then
        mudtProps.IsDirty = True
    End If
    SetContactpersonPhone = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetContactpersonPhone(val)", val
End Function
Public Function SetCell(val As String) As Boolean
Dim strSearchPhone As String
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetCell = True
    If Len(val) > Len(mudtProps.MOBILE) Then
        Err.Raise 384
    End If
    mudtProps.MOBILE = val
    bOK = ValidateObject(val, "CELL")
'    If Len(Trim(mudtProps.MOBILE)) = 0 Then
'        mudtProps.SearchPhone = CleanForSearchField(val)  'Right(StripToNumerics(val), 6)
'    End If
    If bOK Then
        mudtProps.IsDirty = True
    End If
    SetCell = bOK
    Exit Function
End Function

Public Function SetName(val As String) As Boolean
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetName = True
    If Len(val) > Len(mudtProps.Name) Then
        Err.Raise 384
    End If
    mudtProps.Name = val
    bOK = ValidateObject(val, "NAME")
    If bOK Then
        mudtProps.IsDirty = True
'        If bReadyForDuplicateCheck Then
'            LookforDuplicates
'        End If
    End If
    SetName = bOK
    Exit Function
End Function
Public Sub SetBlocked(bBlocked As Boolean)
    mudtProps.Blocked = bBlocked
    SetDirty True
End Sub
Public Sub SetoLineToInvoice(boLineToInvoice As Boolean)
    mudtProps.oLineToInvoice = boLineToInvoice
    SetDirty True
End Sub
Public Sub SetUseQuotedPrice(bUseQuotedPrice As Boolean)
    mudtProps.UseQuotedPrice = bUseQuotedPrice
    SetDirty True
End Sub
Public Sub SetOrderCOmplete(val As Boolean)
    mudtProps.CompleteOrder = val
    SetDirty True
End Sub
Public Property Get Note() As String
    On Error GoTo errHandler
  Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Note"
End Property
Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetNote = True
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    mudtProps.Note = val
    SetNote = ValidateObject(val, "Note")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get VatNumber() As String
    On Error GoTo errHandler
    VatNumber = FNS(mudtProps.VatNumber)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.VATNumber"
End Property
Public Function SetVATNumber(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetVATNumber = True
    If Len(val) > Len(mudtProps.VatNumber) Then
        Err.Raise 384
    End If
    mudtProps.VatNumber = val
    SetVATNumber = ValidateObject(val, "VATNumber")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get Title() As String
    On Error GoTo errHandler
  Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Title"
End Property
'

Public Function SetTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetTitle = True
    If Len(val) > Len(mudtProps.Title) Then
        Err.Raise 384
    End If
    mudtProps.Title = val
    SetTitle = ValidateObject(val, "TITLE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get OurACnoWithClient() As String
    On Error GoTo errHandler
    OurACnoWithClient = FNS(mudtProps.OurACnoWithClient)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.OurACnoWithClient"
End Property
Public Function SetOurACnoWithClient(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetOurACnoWithClient = True
    If Len(val) > Len(mudtProps.OurACnoWithClient) Then
        Err.Raise 384
    End If
    mudtProps.OurACnoWithClient = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Function SetCustomerTypeCasual() As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCustomerTypeCasual = True
    mudtProps.CanBeDeleted = True
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetCustomerTypeCasual"
End Function

Public Property Get DateRecordAddedF() As String
    On Error GoTo errHandler
    If mudtProps.DateRecordAdded = CDate(0) Then
        DateRecordAddedF = "unknown"
    Else
        DateRecordAddedF = Format(mudtProps.DateRecordAdded, "dd/mm/yyyy HH:nn")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.DateRecordAddedF"
End Property
Public Property Get DateRecordLastChangedF() As String
    On Error GoTo errHandler
    If mudtProps.DateLastModified = CDate(0) Then
        DateRecordLastChangedF = "unknown"
    Else
        DateRecordLastChangedF = Format(mudtProps.DateLastModified, "dd/mm/yyyy HH:nn")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.DateRecordLastChangedF"
End Property
Public Property Get Fullname() As String
    On Error GoTo errHandler
    Fullname = Name & IIf(Len(Title) > 0 Or Len(Initials) > 0, ", ", "") & Title & IIf(Len(Title) > 0, " ", "") & Initials
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Fullname"
End Property
Public Property Get NameAndCode(pLen As Integer) As String
    On Error GoTo errHandler
Dim strNameandCode As String
    strNameandCode = Left(Fullname, pLen) & IIf(Len(Fullname) > pLen, "...  ", "  ")
    If AcNo > "" Then
        strNameandCode = strNameandCode & "(" & AcNo & ")"
    End If
    NameAndCode = strNameandCode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.NameAndCode(pLen)", pLen
End Property

Public Property Get Name() As String
    On Error GoTo errHandler
  Name = FNS(mudtProps.Name)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Name"
End Property
Public Property Get MOBILE() As String
    On Error GoTo errHandler
  MOBILE = FNS(mudtProps.MOBILE)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Mobile"
End Property
Public Property Get MOBILEF() As String
    If Len(FNS(mudtProps.MOBILE)) = 10 Then
        MOBILEF = Format(FNS(mudtProps.MOBILE), "\(@@@\)\ @@@\-@@@@")
    Else
        MOBILEF = FNS(mudtProps.MOBILE)
    End If
End Property
Public Property Get ContactpersonPhone() As String
    On Error GoTo errHandler
  ContactpersonPhone = FNS(mudtProps.ContactpersonPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ContactpersonPhone"
End Property
Public Property Get ContactpersonPhoneF() As String
    If Len(FNS(mudtProps.ContactpersonPhone)) = 10 Then
        ContactpersonPhoneF = Format(FNS(mudtProps.ContactpersonPhone), "\(@@@\)\ @@@\-@@@@")
    Else
        ContactpersonPhoneF = FNS(mudtProps.ContactpersonPhone)
    End If
End Property
Public Function SetContactPerson(val As String) As Boolean
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetContactPerson = True
    If Len(val) > Len(mudtProps.ContactPerson) Then
        Err.Raise 384
    End If
    mudtProps.ContactPerson = val
    bOK = ValidateObject(val, "")
    If bOK Then
        mudtProps.IsDirty = True
    End If
    SetContactPerson = bOK
    Exit Function
End Function
Public Property Get ContactPerson() As String
    On Error GoTo errHandler
  ContactPerson = FNS(mudtProps.ContactPerson)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ContactPerson"
End Property

Public Function SetControl(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    mControl = val
    SetControl = True
    SetControl = ValidateObject(val, "CONTROL")
    Exit Function
End Function

Public Property Get SAN() As String
    On Error GoTo errHandler
  SAN = FNS(mudtProps.SAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SAN"
End Property
Public Function SetSAN(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetSAN = True
    If Len(val) > Len(mudtProps.SAN) Then
        Err.Raise 384
    End If
    mudtProps.SAN = UCase(val)
    SetSAN = ValidateObject(val, "SAN")
    mudtProps.IsDirty = True
    Exit Function
End Function




Public Property Get AcNo() As String
    On Error GoTo errHandler
  AcNo = FNS(mudtProps.AcNo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.AcNo"
End Property
Public Function SetAcNO(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetAcNO = True
    If Len(val) > Len(mudtProps.AcNo) Then
        Err.Raise 384
    End If
    mudtProps.AcNo = UCase(val)
    SetAcNO = ValidateObject(val, "ACNO")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get OpenItemOrBalBF() As String
    On Error GoTo errHandler
    OpenItemOrBalBF = FNS(mudtProps.OpenItemOrBalBF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.OpenItemOrBalBF"
End Property
Public Function SetOpenItem(val As Boolean) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OpenItemOrBalBF = IIf(val = True, "O", "B")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get AccAcno() As String
    On Error GoTo errHandler
    AccAcno = FNS(mudtProps.AccAcno)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.AccAcNo"
End Property
Public Function SetAccAcNo(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetAccAcNo = True
    If Len(val) > Len(mudtProps.AccAcno) Then
        Err.Raise 384
    End If
    mudtProps.AccAcno = Trim(UCase(val))
    SetAccAcNo = ValidateObject(val, "ACCACNO")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IsNew"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IsEditing"
End Property
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IsValid"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = mudtProps.IsDirty Or Me.Addresses.IsDirty Or Me.InterestGroups.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IsDirty"
End Property
Public Function Validate(pMsg As String) As Boolean
    On Error GoTo errHandler
    pMsg = mobjValid.AllBrokenRules
    Validate = (pMsg = "")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Validate(pMsg)", pMsg
End Function
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
Dim lngTmp As Long
    bValid = True
    If Not (Len(FNS(mudtProps.Name)) > 1) Then  'Or Len(FNS(mudtProps.Name)) = 0
        mobjValid.BreakRule "NAME", True
        If pFld = "NAME" Then bValid = False
    Else
        mobjValid.BreakRule "NAME", False
    End If
    If Not (Len(FNS(mudtProps.Title)) > 1 Or Len(FNS(mudtProps.Title)) = 0) Then
        mobjValid.BreakRule "TITLE", True
        If pFld = "TITLE" Then bValid = False
    Else
        mobjValid.BreakRule "TITLE", False
    End If
    If Me.CustomerTypes.IsALoyaltyMember Or Me.CustomerTypeID = oPC.Configuration.LoyaltyClubTypeID Then
        If Not IsNumeric(FNS(mudtProps.AcNo)) Then
            mobjValid.BreakRule "ACNO", True
            If pFld = "ACNO" Then bValid = False
        Else
            ConvertToLng FNS(mudtProps.AcNo), lngTmp
            If lngTmp > oPC.Configuration.LoyaltyNumberingEndAt Or lngTmp < oPC.Configuration.LoyaltyNumberingStartAt Then
                mobjValid.BreakRule "ACNO", True
                If pFld = "ACNO" Then bValid = False
            Else
                mobjValid.BreakRule "ACNO", False
            End If
          'MsgBox "Fix here"
        End If
    Else
        If (Not Len(FNS(mudtProps.AcNo)) > 2) And (oPC.GetProperty("EnforceCustomerAcno") = "TRUE") Then
            mobjValid.BreakRule "ACNO", True
            If pFld = "ACNO" Then bValid = False
        Else
            mobjValid.BreakRule "ACNO", False
        End If
    
    End If
    If Not (Len(FNS(mudtProps.Note)) > 2 Or Len(FNS(mudtProps.Note)) = 0) Then
        mobjValid.BreakRule "NOTE", True
        If pFld = "NOTE" Then bValid = False
    Else
        mobjValid.BreakRule "NOTE", False
    End If
    If mudtProps.DefaultDiscount > 100 Then
        mobjValid.BreakRule "DISCOUNT", True
        If pFld = "DISCOUNT" Then bValid = False
    Else
        mobjValid.BreakRule "DISCOUNT", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ValidateObject(pVal,pFld)", Array(pVal, pFld)
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
    ErrorIn "a_Customer.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Name too short", "NAME"
    colClassors.Add "Title too short", "TITLE"
    colClassors.Add "Initials too short", "INITIALS"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "Phone missing or too short", "PHONE"
    If oPC.Configuration.LoyaltyNumberingEndAt > 0 Then
        colClassors.Add "Account number invalid, either too short or not in loyalty number range. (" & CStr(oPC.Configuration.LoyaltyNumberingStartAt) & " to " & CStr(oPC.Configuration.LoyaltyNumberingEndAt) & ") or in Loyalty number range but not a loyalty customer", "ACNO"
    Else
        colClassors.Add "Account number invalid.", "ACNO"
    End If
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "Missing duplicate checking field", "CONTROL"
    colClassors.Add "V.A.T. number missing", "VAT"
'    colClassors.Add "Both (phone or Email) and name must be supplied ", "MINIMUM"
'    colClassors.Add "Can't have an interest group without and email address", "IG_EMAIL"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.LoadClassorsCollection"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetDirty(pVal)", pVal
End Sub

Public Sub ExportStatement(dteFrom As Date, dteTo As Date, dteStatementDate As Date, Optional pFilename As String)
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim oTRs As c_DebtorsTransPerTP
Dim i As Long
Dim j As Long
Dim dblBal As Double
Dim strXML As String

'If Me.AcNo = "KLV121" Then MsgBox "HERE"

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "STATEMENT_DOC"
        .chCreate "DestinationAddress"
                If Me.DelToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Me.DelToAddress.EMail
                End If
        .elCreateSibling "TemplateName"
        .elText = "STATEMENT.XSL"
        .elCreateSibling "StoreCode", True
            .elText = oPC.Configuration.DefaultStore.code
        .elCreateSibling "StatementDate", True
            .elText = Format(dteStatementDate, "DD/mm/yyyy")
        .elCreateSibling "Sender", True
            .elText = oPC.Configuration.DefaultCompany.CompanyName                  'oPC.Configuration.DefaultCompany.CompanyName
        .elCreateSibling "SenderAddress", True
            .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
        .elCreateSibling "CustomerName", True
            .elText = Me.Name
        .elCreateSibling "CustomerNameandAddress", True
        If Not Me.BillTOAddress Is Nothing Then
            .elText = Me.BillTOAddress.AddressMailing
        End If
        .elCreateSibling "BankingDetails", True
            .elText = oPC.Configuration.DefaultCompany.BankDetails
        .elCreateSibling "OurVATNumber", True
            .elText = oPC.Configuration.DefaultCompany.VatNumber
            
        .elCreateSibling "CustomerVATNumber", True
        If Me.VatNumber > "" Then
            .elText = "Customer's VAT number: " & Me.VatNumber
        Else
            .elText = ""
        End If
        .elCreateSibling "ACNO", True
            .elText = Me.AcNo
        .elCreateSibling "PaymentStyle", True
            .elText = Me.PaymentStyle
            
            
        Set oTRs = Nothing
        Set oTRs = New c_DebtorsTransPerTP
        oTRs.LoadOpenItem Me.ID, dteFrom, dteTo
        i = 1
        Do While i <= oTRs.Count
            .elCreateSibling "DetailLine_OI", True
            .chCreate "dbDoc"
                .elText = oTRs.Item(i).dbDoc
            .elCreateSibling "dbDocType", True
                .elText = oTRs.Item(i).dbDocType
            .elCreateSibling "dbDate", True
                .elText = oTRs.Item(i).dbDate
            .elCreateSibling "dbAmt", True
                .elText = oTRs.Item(i).dbAmtF
            .elCreateSibling "crDoc"
                .elText = oTRs.Item(i).crDoc
            .elCreateSibling "crDocType", True
                .elText = oTRs.Item(i).crDocType
            .elCreateSibling "crDate", True
                .elText = oTRs.Item(i).crDate
            .elCreateSibling "crAmt", True
                .elText = oTRs.Item(i).crAmtF
            .elCreateSibling "crTotal", True
                .elText = oTRs.Item(i).crTotalF
            .elCreateSibling "balance", True
                .elText = oTRs.Item(i).Balance_OIF
            i = i + 1
            .navUP
        Loop
            .elCreateSibling "DetailLine_OI", True
            .chCreate "dbDoc"
                .elText = "Total"
            .elCreateSibling "dbDocType", True
                .elText = ""
            .elCreateSibling "dbDate", True
                .elText = ""
            .elCreateSibling "dbAmt", True
                .elText = ""
            .elCreateSibling "crDoc"
                .elText = ""
            .elCreateSibling "crDocType", True
                .elText = ""
            .elCreateSibling "crDate", True
                .elText = ""
            .elCreateSibling "crAmt", True
                .elText = ""
            .elCreateSibling "crTotal", True
                .elText = ""
            .elCreateSibling "balance", True
                If oTRs.Count > 0 Then
                    .elText = oTRs.Item(i - 1).PayableAmount
                Else
                    .elText = ""
                End If
            .navUP
            Set oTRs = Nothing
            .elCreateSibling "Balances", True
            .chCreate "TotalBalance"
                .elText = Me.BalanceF
            .elCreateSibling "Current"
                .elText = Me.BalanceCurF
            .elCreateSibling "Thirty"
                .elText = Me.Balance30F
            .elCreateSibling "Sixty"
                .elText = Me.Balance60F
            .elCreateSibling "Ninety"
                .elText = Me.Balance90F
            .elCreateSibling "OneTwenty"
                .elText = Me.Balance120F
            .navUP
            .elCreateSibling "BFBal", True
                .elText = Me.BalBFF
        .elCreateSibling "Memo", True
            .elText = oPC.Configuration.StatementText
        .elCreateSibling "CompanyRegistration", True
            .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
        .elCreateSibling "VATNumber", True
            .elText = oPC.Configuration.DefaultCompany.VatNumber
    End With
'FINALLY PRODUCE THE .XML FILE
    If pFilename > "" Then
        strXML = oPC.SharedFolderRoot & "\Statements\" & pFilename
    Else
        strXML = oPC.SharedFolderRoot & "\Statements\ST_" & IIf(Me.AcNo > "", Me.AcNo, CStr(Me.ID)) & ".xml"
    End If
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ExportStatement"
End Sub

'Public Function SetDispatchMethod(val As String) As Boolean
'Dim bOK As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    bOK = IIf(UCase(val) = "E" Or UCase(val) = "M" Or UCase(val) = "P", True, False)
'    If bOK Then
'        mudtProps.DispatchMethod = val
'    End If
'    SetDispatchMethod = bOK
'End Function
'================================================================================
Public Property Get SalesOrderTemplateName() As String
    On Error GoTo errHandler
    SalesOrderTemplateName = FNS(mudtProps.SalesOrderTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SalesOrderTemplateName"
End Property
Public Function SetSalesOrderTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetSalesOrderTemplateName = True
    If Len(val) > Len(mudtProps.SalesOrderTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.SalesOrderTemplateName = val
    SetSalesOrderTemplateName = ValidateObject(val, "SalesOrderTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get QuotationTemplateName() As String
    On Error GoTo errHandler
    QuotationTemplateName = FNS(mudtProps.QuotationTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.QuotationTemplateName"
End Property
Public Function SetQuotationTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetQuotationTemplateName = True
    If Len(val) > Len(mudtProps.QuotationTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.QuotationTemplateName = val
    SetQuotationTemplateName = ValidateObject(val, "QuotationTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get ApproTemplateName() As String
    On Error GoTo errHandler
    ApproTemplateName = FNS(mudtProps.ApproTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ApproTemplateName"
End Property
Public Function SetApproTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetApproTemplateName = True
    If Len(val) > Len(mudtProps.ApproTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.ApproTemplateName = val
    SetApproTemplateName = ValidateObject(val, "ApproTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get ApproReturnTemplateName() As String
    On Error GoTo errHandler
    ApproReturnTemplateName = FNS(mudtProps.ApproReturnTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ApproReturnTemplateName"
End Property
Public Function SetApproReturnTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetApproReturnTemplateName = True
    If Len(val) > Len(mudtProps.ApproReturnTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.ApproReturnTemplateName = val
    SetApproReturnTemplateName = ValidateObject(val, "ApproReturnTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get InvoiceTemplateName() As String
    On Error GoTo errHandler
    InvoiceTemplateName = FNS(mudtProps.InvoiceTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.InvoiceTemplateName"
End Property
Public Function SetInvoiceTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetInvoiceTemplateName = True
    If Len(val) > Len(mudtProps.InvoiceTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.InvoiceTemplateName = val
    SetInvoiceTemplateName = ValidateObject(val, "InvoiceTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function


Public Property Get GDNTemplateName() As String
    On Error GoTo errHandler
    GDNTemplateName = FNS(mudtProps.GDNTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GDNTemplateName"
End Property
Public Function SetGDNTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetGDNTemplateName = True
    If Len(val) > Len(mudtProps.GDNTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.GDNTemplateName = val
    SetGDNTemplateName = ValidateObject(val, "GDNTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function


Public Property Get CreditNoteTemplateName() As String
    On Error GoTo errHandler
    CreditNoteTemplateName = FNS(mudtProps.CreditNoteTemplateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CreditNoteTemplateName"
End Property
Public Function SetCreditNoteTemplateName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCreditNoteTemplateName = True
    If Len(val) > Len(mudtProps.CreditNoteTemplateName) Then
        Err.Raise 384
    End If
    mudtProps.CreditNoteTemplateName = val
    SetCreditNoteTemplateName = ValidateObject(val, "CreditNoteTemplateName")
    mudtProps.IsDirty = True
    Exit Function
End Function

