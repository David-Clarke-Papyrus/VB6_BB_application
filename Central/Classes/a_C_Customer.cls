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
Attribute VB_Exposed = True
Option Explicit

Event Valid(strMsg As String)
Event DefaultAddressChanged()
Event Dirty(bYesNO As Boolean)
Event PossibleDuplicates(Duplicates As c_C_Customer)
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

Private tlTitles As z_TextList
Private tlInterestGroupsActive As z_TextList
Private tlInterestGroupsAll As z_TextList
'Private tlCustomerTypesActive As z_TextList
'Private tlCustomerTypesAll As z_TextList
Private tlCustomerTypes As z_TextList
'Private cIGs As c_IGs
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassErrors As Collection
Private mApproAddressIdx As String
Private mBilltoAddressIdx As String
Private mDelToAddressIdx As String
Private mOrderToAddressIdx As String
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim mControl As String
Dim bReadyForDuplicateCheck As Boolean
Public Property Get InterestGroups() As ch_IG
    Set InterestGroups = mcolIG
End Property
Public Property Get CustomerTypes() As ch_IG
    Set CustomerTypes = mcolCC
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
    mobjValid.GetSTatus
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
Dim oADd As a_Address
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
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Dim objPersist As a_Customer_P
    lngResult = 0
    
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
EXIT_Handler:
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    lngResult = 101
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ApplyEdit(lngResult)", lngResult, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Sub

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
   ' Set tlCustomerTypesActive = New z_TextList
   ' Set tlCustomerTypesAll = New z_TextList
    Set tlCustomerTypes = New z_TextList
'    Set cIGs = New c_IGs
    tlInterestGroupsActive.Load ltInterestGroupActive
    tlInterestGroupsAll.Load ltInterestGroupAll
    tlCustomerTypes.Load ltCustomerTypeAll
 '   tlCustomerTypesActive.Load ltCustomerTypeActive
 '   tlCustomerTypesAll.Load ltCustomerTypeAll
    mudtProps.CustomerTypeID = oPC.Configuration.LoyaltyClubTypeID
    mudtProps.IsNew = True
    Set colClassErrors = New Collection
    LoadClassErrorsCollection
    mBilltoAddressIdx = 1
    mobjValid.BreakRule "STORE", True
    mobjValid.RuleBroken "NAME", True
    VATable = True
    bReadyForDuplicateCheck = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Class_Initialize"
End Sub
Public Sub InitializeNewCustomer()
Dim oADd As a_Address
    Set oADd = Addresses.Add
    oADd.BeginEdit
    oADd.SetDescription "Default"
    oADd.ApplyEdit
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
  '  Set tlCustomerTypesActive = Nothing
    Set tlInterestGroupsAll = Nothing
 '   Set tlCustomerTypesAll = Nothing
    Set tlCustomerTypes = Nothing
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Class_Terminate"
End Sub


Private Sub mobjValid_BrokenRule(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mobjValid_BrokenRule(errors)", errors
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.mobjValid_RuleUnbroken(errors)", errors
End Sub

'Private Sub mobjValid_ThisRuleUnbroken(Thiserror As String)
'    On Error GoTo errHandler
'    If Thiserror = "MINIMUM" Then 'we can check for duplicates
'        LookforDuplicates
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.mobjValid_ThisRuleUnbroken(errors)", errors
'End Sub
Public Sub LookforDuplicates()
Dim oDup As New c_C_Customer
Dim bFound As Boolean
Dim tmpCust As a_Customer
    bFound = False
    If Me.SearchPhone > "" Then
        oDup.LoadDuplicates bFound, Me.SearchPhone, Me.ID
        If bFound Then
            RaiseEvent PossibleDuplicates(oDup)
        End If
    End If
End Sub
Public Function CustomerIndexClashes() As Boolean
Dim oSM As New z_StockManager
    
    If Me.IsNew Then
        CustomerIndexClashes = oSM.CustomerIndexClashes(0, Me.AcNo)
    Else
        CustomerIndexClashes = oSM.CustomerIndexClashes(Me.ID, Me.AcNo)
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
    ErrorIn "a_Customer.GetSTatus"
End Function

Public Property Get InterestGroupsAll_tl() As z_TextList
    Set InterestGroupsAll_tl = tlInterestGroupsAll
End Property
Public Property Get InterestGroupsActive_tl() As z_TextList
    Set InterestGroupsActive_tl = tlInterestGroupsActive
End Property
Public Property Get CustomerTypes_tl() As z_TextList
    Set CustomerTypes_tl = tlCustomerTypes
End Property
'Public Property Get CustomerTypesALL_tl() As z_TextList
'    Set CustomerTypesALL_tl = tlCustomerTypesAll
'End Property
'Public Property Get CustomerTypesActive_tl() As z_TextList
'    Set CustomerTypesActive_tl = tlCustomerTypesActive
'End Property
Public Function Load(Optional CustomerID As Long = 0, Optional pACNO As String = "", Optional pDefaultPhone As String = "") As Boolean
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
'Public Function ReloadIGs()
'    On Error GoTo errHandler
'    Set cIGs = Nothing
'    Set cIGs = New c_IGs
'    cIGs.Load mudtProps.CustID
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.ReloadIGs"
'End Function


Public Sub SetApproAddressidx(val As String)
Dim oADd As a_Address
    On Error GoTo errHandler
    mApproAddressIdx = val
    Set oApproAddress = Nothing
    Set oApproAddress = Me.Addresses(mApproAddressIdx)
    For Each oADd In Addresses
        oADd.Appro = False
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
Dim oADd As a_Address

    On Error GoTo errHandler
    mBilltoAddressIdx = val
    Set oBillToAddress = Nothing
    Set oBillToAddress = Addresses(mBilltoAddressIdx)
    For Each oADd In Addresses
        oADd.BillTo = False
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
Dim oADd As a_Address
    On Error GoTo errHandler
    mDelToAddressIdx = val
    Set oDelToAddress = Nothing
    Set oDelToAddress = Me.Addresses(mDelToAddressIdx)
    For Each oADd In Addresses
        oADd.DelTo = False
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
Dim oADd As a_Address
    On Error GoTo errHandler
    mOrderToAddressIdx = val
    Set oOrderToAddress = Nothing
    Set oOrderToAddress = Me.Addresses(mOrderToAddressIdx)
    For Each oADd In Addresses
        oADd.OrderTo = False
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
'    CustomerTypeShortname = tlCustomerTypes.Item(mudtProps.CustomerTypeID)
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
'Public Function AddIG(pIGID As Long)
'    On Error GoTo errHandler
'Dim oCustPersist As a_Customer_P
'    Set oCustPersist = New a_Customer_P
'    oCustPersist.AddIG Me.ID, pIGID
'    Set oCustPersist = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.AddIG(pIGID)", pIGID
'End Function
'Public Function RemoveIG(pIGID As Long)
'    On Error GoTo errHandler
'Dim oCustPersist As a_Customer_P
'    Set oCustPersist = New a_Customer_P
'    oCustPersist.RemoveIG Me.ID, pIGID
'    Set oCustPersist = Nothing
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.RemoveIG(pIGID)", pIGID
'End Function
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
        bOK = ValidateObject(val, "DEFAULTDISCOUNT")
    End If
    SetDefaultDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetDefaultDiscount(val)", val
End Function
Public Property Get Titles() As z_TextList
    Set Titles = tlTitles
End Property

'Public Function OKForDeletion(pXASafe As XArrayDB, pXAUnSafe As XArrayDB, pDPTP As c_DocsPerTP) As Boolean
'    On Error GoTo errHandler
'Dim oACt As New c_DocsPerTP
'Dim i As Long
'
'    OKForDeletion = True
'    Set XA = Nothing
'    Set XA = pXASafe
'    oACt.Load Me.ID, False
'    XA.ReDim 1, oACt.Count, 1, 7
'    If oACt.Count > 0 Then
'        OKForDeletion = False
'        For i = 1 To oACt.Count
'            XA(i, 1) = oACt.Item(i).DocDateF
'            XA(i, 2) = oACt.Item(i).DocCode
'            XA(i, 3) = oACt.Item(i).DocTypeF
'            XA(i, 4) = oACt.Item(i).DocStatus
'          '  XA(i, 5) = oACt.Item(i).DocCode
'            XA(i, 6) = oACt.Item(i).DOCID
'            XA(i, 7) = oACt.Item(i).DocDateForSort
'        Next
'        Set pDPTP = oACt
'    End If
'
'    Set XB = Nothing
'    Set XB = pXAUnSafe
'    oACt.Load Me.ID, True
'    XB.ReDim 1, oACt.Count, 1, 7
'    If oACt.Count > 0 Then
'        OKForDeletion = False
'        For i = 1 To oACt.Count
'            XB(i, 1) = oACt.Item(i).DocDateF
'            XB(i, 2) = oACt.Item(i).DocCode
'            XB(i, 3) = oACt.Item(i).DocType
'            XB(i, 4) = oACt.Item(i).DocStatus
'          '  XA(i, 5) = oACt.Item(i).DocCode
'            XB(i, 6) = oACt.Item(i).DOCID
'            XB(i, 7) = oACt.Item(i).DocDateForSort
'        Next
'        Set pDPTP = oACt
'        OKForDeletion = False
'
'    End If
'
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.OKForDeletion(pXASafe,pXAUnSafe,pDPTP)", Array(pXASafe, pXAUnSafe, pDPTP)
'End Function


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
    On Error GoTo errHandler
    If Not BillTOAddress Is Nothing Then
        Phonef = BillTOAddress.Phonef
    Else
        Phonef = FNS(mudtProps.Phone)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Phone"
End Property
'Public Property Get Cell() As String
'    On Error GoTo errHandler
'    If Not BillTOAddress Is Nothing Then
'        Phone = BillTOAddress.Phone
'    Else
'        Phone = FNS(mudtProps.Phone)
'    End If
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer.Phone"
'End Property

Public Property Get SearchPhone() As String
    On Error GoTo errHandler
    SearchPhone = FNS(mudtProps.SearchPhone)
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
Dim strSearchPhone As String
'This is used to store either phone or email
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetPhone = True
    If Len(val) > Len(mudtProps.Phone) Then
        Err.Raise 384
    End If
    mudtProps.Phone = val
    bOK = ValidateObject(val, "PHONE")
    mudtProps.SearchPhone = CleanForSearchField(val)  'Right(StripToNumerics(val), 6)
    If bOK Then
        mudtProps.IsDirty = True
'        If bReadyForDuplicateCheck Then
'            LookforDuplicates
'        End If
    End If
    SetPhone = bOK
    Exit Function
End Function
Public Function SetMobile(val As String) As Boolean
Dim strSearchPhone As String
Dim bOK As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetMobile = True
    If Len(val) > Len(mudtProps.Mobile) Then
        Err.Raise 384
    End If
    mudtProps.Mobile = val
    bOK = ValidateObject(val, "Mobile")
    If Len(Trim(mudtProps.Phone)) = 0 Then
        mudtProps.SearchPhone = CleanForSearchField(val)  'Right(StripToNumerics(val), 6)
    End If
    If bOK Then
        mudtProps.IsDirty = True
    End If
    SetMobile = bOK
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
Public Property Get StoreName() As String
    On Error GoTo errHandler
    StoreName = FNS(mudtProps.StoreName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Storename"
End Property
Public Property Get StoreID() As Long
    On Error GoTo errHandler
    StoreID = FNN(mudtProps.StoreID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.StoreID"
End Property
Public Function SetStoreID(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    SetStoreID = True
    mudtProps.StoreID = val
    If val > 0 Then
        mobjValid.BreakRule "STORE", False
    Else
        mobjValid.BreakRule "STORE", True
    End If
    mobjValid.GetSTatus

    mudtProps.IsDirty = True
End Function
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
Public Property Get IDNUM() As String
    On Error GoTo errHandler
  IDNUM = FNS(mudtProps.IDNUM)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IDNUM"
End Property
Public Function SetIDNUM(val As String) As Boolean
Dim flgValid As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetIDNUM = True
    If Len(val) > Len(mudtProps.IDNUM) Then
        Err.Raise 384
    End If
    mudtProps.IDNUM = val
    SetIDNUM = ValidateObject(val, "IDNUM")
    mudtProps.IsDirty = True
    Exit Function
End Function
'Public Function SetOrginatingStore(val As Long) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetOrginatingStore = True
'    mudtProps.StoreID = val
'    mudtProps.IsDirty = True
'    Exit Function
'End Function

Public Property Get Title() As String
    On Error GoTo errHandler
  Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Title"
End Property
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
        DateRecordAddedF = Format(mudtProps.DateRecordAdded, "dd/mm/yyyy")
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
        DateRecordLastChangedF = Format(mudtProps.DateLastModified, "dd/mm/yyyy")
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
Public Property Get Mobile() As String
    On Error GoTo errHandler
  Mobile = FNS(mudtProps.Mobile)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Name"
End Property
Public Property Get MobileF() As String
    If Len(FNS(mudtProps.Mobile)) = 10 Then
        MobileF = Format(FNS(mudtProps.Mobile), "\(@@@\)\ @@@\-@@@@")
    Else
        MobileF = FNS(mudtProps.Mobile)
    End If
End Property

Public Function SetControl(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    mControl = val
    SetControl = True
    SetControl = ValidateObject(val, "CONTROL")
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
    On Error Resume Next
    If Len(val) > Len(mudtProps.AcNo) Then
        Err.Raise 384
    End If
    mudtProps.AcNo = UCase(val)
    SetAcNO = ValidateObject(val, "ACNO")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get IsLoyaltyClubMember() As Boolean
    On Error GoTo errHandler
    IsLoyaltyClubMember = mcolCC.IsALoyaltyMember
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.IsDeleted"
End Property
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
    bValid = True
'    If Not (Len(FNS(mudtProps.Name)) > 1 And Len(FNS(mudtProps.Phone)) > 5) Then
'        mobjValid.BreakRule "MINIMUM", True
'        If pFld = "MINIMUM" Then bValid = False
'    Else
'        mobjValid.BreakRule "MINIMUM", False
'        If pFld = "NAME" Or pFld = "PHONE" Then bReadyForDuplicateCheck = True
'    End If
    If Not (Len(FNS(mudtProps.Name)) > 2 Or Len(FNS(mudtProps.Name)) = 0) Then
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
    If Not (Len(FNS(mudtProps.AcNo)) > 2 Or Len(FNS(mudtProps.AcNo)) = 0) Then
        mobjValid.BreakRule "ACNO", True
        If pFld = "ACNO" Then bValid = False
    Else
        mobjValid.BreakRule "ACNO", False
    End If
    If Not (Len(FNS(mudtProps.Note)) > 2 Or Len(FNS(mudtProps.Note)) = 0) Then
        mobjValid.BreakRule "NOTE", True
        If pFld = "NOTE" Then bValid = False
    Else
        mobjValid.BreakRule "NOTE", False
    End If
    If FNN(mudtProps.StoreID) = 0 Then
        mobjValid.BreakRule "STORE", True
        If pFld = "STORE" Then bValid = False
    Else
        mobjValid.BreakRule "STORE", False
    End If
    If (Not ValidIDNum(Trim(mudtProps.IDNUM)) And Trim(mudtProps.IDNUM) > "") Then
        mobjValid.BreakRule "IDNUM", True
        If pFld = "IDNUM" Then bValid = False
    Else
        mobjValid.BreakRule "IDNUM", False
    End If
    mobjValid.GetSTatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ValidateObject(pVal,pFld)", Array(pVal, pFld)
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
    ErrorIn "a_Customer.TranslateErrors(pRawErrors)", pRawErrors
End Function
Private Sub LoadClassErrorsCollection()
    On Error GoTo errHandler
    Set colClassErrors = New Collection
    colClassErrors.Add "Name too short", "NAME"
    colClassErrors.Add "Title too short", "TITLE"
    colClassErrors.Add "Invalid I.D.number", "IDNUM"
    colClassErrors.Add "Initials too short", "INITIALS"
    colClassErrors.Add "Phone missing or too short", "PHONE"
    colClassErrors.Add "Account number too short", "ACNO"
    colClassErrors.Add "Note too short", "NOTE"
    colClassErrors.Add "Missing duplicate checking field", "CONTROL"
    colClassErrors.Add "Both (phone or Email) and name must be supplied ", "MINIMUM"
    colClassErrors.Add "Missing source store", "STORE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.LoadClassErrorsCollection"
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
Public Property Let ExcludeFromSales(val As Boolean)
    mudtProps.ExcludeFromSales = FNB(val)
End Property
Public Property Get ExcludeFromSales() As Boolean
    ExcludeFromSales = FNB(mudtProps.ExcludeFromSales)
End Property
Function ValidIDNum(pIn) As Boolean
Dim a As Long
Dim b As Long
Dim c As Long

Dim tmpB As String
Dim ichkDigit As Integer

Dim i As Integer
    If IsNull(pIn) Then
        ValidIDNum = False
        GoTo EXIT_Handler
    End If
    If Not IsNumeric(pIn) Then
        ValidIDNum = False
        GoTo EXIT_Handler
    End If
    If Len(pIn) <> 13 Then
        ValidIDNum = False
        GoTo EXIT_Handler
    End If
    a = 0
    For i = 1 To 11 Step 2
        a = a + CInt(Mid(pIn, i, 1))
    Next i
    tmpB = ""
    For i = 2 To 12 Step 2
        tmpB = tmpB & Mid(pIn, i, 1)
    Next i
    c = 2 * (CLng(tmpB))
    tmpB = CStr(c)
    b = 0
    For i = 1 To Len(tmpB)
        b = b + CInt(Mid(tmpB, i, 1))
    Next i
    ichkDigit = 10 - Right(CStr(a + b), 1)
    If ichkDigit = 10 Then ichkDigit = 0
    If ichkDigit = CInt(Right(pIn, 1)) Then
        ValidIDNum = True
    Else
        ValidIDNum = False
    End If
    
EXIT_Handler:
    Exit Function
End Function

