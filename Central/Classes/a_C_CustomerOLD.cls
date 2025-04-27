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
Private mudtProps As CustomerProps
Private mcolStack As Collection
Private WithEvents mcolAddresses As ch_Addresses
Attribute mcolAddresses.VB_VarHelpID = -1
Private oApproAddress As a_Address
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oOrderToAddress As a_Address

Private tlTitles As z_TextList
Private tlInterestGroups As z_TextList
Private tlCustomerTypes As z_TextList
Private cIGs As c_IGs
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassErrors As Collection
Private mApproAddressIdx As Integer
Private mBilltoAddressIdx As Integer
Private mDelToAddressIdx As Integer
Private mOrderToAddressIdx As Integer
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim mControl As String

Private Sub mcolAddresses_FirstAddressAdded()
    On Error GoTo errHandler
    SetBillToAddressidx 1
    SetApproAddressidx 1
    SetDelToAddressidx 1
    SetOrderToAddressidx 1
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
Public Property Get IGs() As c_IGs
    Set IGs = cIGs
End Property
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "Addresses", mcolAddresses.GetSuperState
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

  Dim objPersist As a_Customer_P
  lngResult = 0
  oPC.CO.BeginTrans
  If mcolStack.Count = 0 Then Err.Raise 445
  If Not oBillToAddress Is Nothing Then
    Me.SetPhone oBillToAddress.Phone
  End If
  mcolAddresses.ApplyEdit
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
        oPC.CO.RollbackTrans
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
  oPC.CO.CommitTrans
EXIT_Handler:
    Exit Sub
   ' Resume
'ERR_Handler:
'    oPC.CO.RollbackTrans
'    lngResult = 101
'    MsgBox "Unknown Error" & Error
'    GoTo EXIT_Handler
'    Exit Sub
errHandler:
    ErrPreserve
    lngResult = 101
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ApplyEdit(lngResult)", lngResult, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.CO
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mcolAddresses = New ch_Addresses
    Set tlTitles = New z_TextList
    Set tlInterestGroups = New z_TextList
    Set tlCustomerTypes = New z_TextList
    Set cIGs = New c_IGs
    tlInterestGroups.Load ltDictionary, , dtInterestGroup, "<none>"
    tlCustomerTypes.Load ltDictionary, , dtCustomerType, "<unknown>"
    mudtProps.CustomerTypeID = tlCustomerTypes.Key("*Casual*")
    mudtProps.IsNew = True
    Set colClassErrors = New Collection
    LoadClassErrorsCollection
    mBilltoAddressIdx = 1
    mobjValid.RuleBroken "CONTROL", True
    mobjValid.RuleBroken "NAME", True
    mobjValid.RuleBroken "PHONE", True
    VATable = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set mcolAddresses = Nothing
    Set tlTitles = Nothing
    Set cIGs = Nothing
    Set tlInterestGroups = Nothing
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

Public Property Get InterestGroups_tl() As z_TextList
    Set InterestGroups_tl = tlInterestGroups
End Property
Public Property Get CustomerTypes_tl() As z_TextList
    Set CustomerTypes_tl = tlCustomerTypes
End Property
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
        cIGs.Load mudtProps.CustID
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
Public Function ReloadIGs()
    On Error GoTo errHandler
    Set cIGs = Nothing
    Set cIGs = New c_IGs
    cIGs.Load mudtProps.CustID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ReloadIGs"
End Function


Public Sub SetApproAddressidx(val As Integer)
    On Error GoTo errHandler
    mApproAddressIdx = val
    Set oApproAddress = Nothing
    Set oApproAddress = Me.Addresses(mApproAddressIdx)
    Addresses.SetAppro (mApproAddressIdx)
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

Public Sub SetBillToAddressidx(val As Integer)
    On Error GoTo errHandler
    mBilltoAddressIdx = val
    Set oBillToAddress = Nothing
    Set oBillToAddress = Me.Addresses(mBilltoAddressIdx)
    Addresses.SetBillTo (mBilltoAddressIdx)
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

Public Sub SetDelToAddressidx(val As Integer)
    On Error GoTo errHandler
    mDelToAddressIdx = val
    Set oDelToAddress = Nothing
    Set oDelToAddress = Me.Addresses(mDelToAddressIdx)
    Addresses.SetDelTo (mDelToAddressIdx)
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

Public Sub SetOrderToAddressidx(val As Integer)
    On Error GoTo errHandler
    mOrderToAddressIdx = val
    Set oOrderToAddress = Nothing
    Set oOrderToAddress = Me.Addresses(mOrderToAddressIdx)
    Addresses.SetOrderTo (mOrderToAddressIdx)
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
    CustomerTypeShortname = tlCustomerTypes.Item(mudtProps.CustomerTypeID)
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
Public Function AddIG(pIGID As Long)
    On Error GoTo errHandler
Dim oCustPersist As a_Customer_P
    Set oCustPersist = New a_Customer_P
    oCustPersist.AddIG Me.ID, pIGID
    Set oCustPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.AddIG(pIGID)", pIGID
End Function
Public Function RemoveIG(pIGID As Long)
    On Error GoTo errHandler
Dim oCustPersist As a_Customer_P
    Set oCustPersist = New a_Customer_P
    oCustPersist.RemoveIG Me.ID, pIGID
    Set oCustPersist = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.RemoveIG(pIGID)", pIGID
End Function
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

Public Function OKForDeletion(pXASafe As XArrayDB, pXAUnSafe As XArrayDB, pDPTP As c_SalesPerTP) As Boolean
    On Error GoTo errHandler
Dim oACt As New c_SalesPerTP
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
            XA(i, 2) = oACt.Item(i).DocCode
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
            XB(i, 2) = oACt.Item(i).DocCode
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

Public Property Get CanBeDeleted() As Boolean
    On Error GoTo errHandler
  CanBeDeleted = Trim$(mudtProps.CanBeDeleted)
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
  Initials = Trim$(mudtProps.Initials)
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
Public Property Get SearchPhone() As String
    On Error GoTo errHandler
  SearchPhone = FNS(mudtProps.SearchPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SearchPhone"
End Property

Public Function SetPhone(val As String) As Boolean
Dim strSearchPhone As String
    If mcolStack.Count = 0 Then Err.Raise 383
    SetPhone = True
  '  val = StripToNumerics(val)
    If Len(val) > Len(mudtProps.Phone) Then
        Err.Raise 384
    End If
    mudtProps.Phone = val
    mudtProps.SearchPhone = Right(StripToNumerics(val), 6)
'    mudtProps.Phone = val
    SetPhone = ValidateObject(val, "PHONE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get Note() As String
    On Error GoTo errHandler
  Note = Trim$(mudtProps.Note)
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

Public Property Get Title() As String
    On Error GoTo errHandler
  Title = Trim$(mudtProps.Title)
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
    mudtProps.CustomerTypeID = tlCustomerTypes.Key("*Casual*")
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
  Name = Trim$(mudtProps.Name)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Name"
End Property
Public Function SetName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetName = True
    If Len(val) > Len(mudtProps.Name) Then
        Err.Raise 384
    End If
    mudtProps.Name = val
    SetName = ValidateObject(val, "NAME")
    mudtProps.IsDirty = True
    Exit Function
End Function
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
    If Len(val) > Len(mudtProps.Title) Then
        Err.Raise 384
    End If
    mudtProps.AcNo = UCase(val)
    SetAcNO = ValidateObject(val, "ACNO")
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
    IsDirty = mudtProps.IsDirty
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
'    If Not (Len(Trim(mudtProps.Initials)) > 1 Or Len(Trim(mudtProps.Initials)) = 0) Then
'        mobjValid.BreakRule "Initials", True
'        If pFld = "INITIALS" Then bValid = False
'    Else
'        mobjValid.BreakRule "INITIALS", False
'    End If
'    If Not (Len(Trim(mudtProps.DefaultDiscount)) > 1 Or Len(Trim(mudtProps.Initials)) = 0) Then
'        mobjValid.BreakRule "Initials", True
'        If pFld = "INITIALS" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "INITIALS", False
'    End If
    If Len(mControl) < 6 Then
        mobjValid.BreakRule "CONTROL", True
        If pFld = "CONTROL" Then bValid = False
    Else
        mobjValid.BreakRule "CONTROL", False
    End If
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
    If Not Len(FNS(mudtProps.Phone)) > 6 Then
        mobjValid.BreakRule "PHONE", True
        If pFld = "PHONE" Then bValid = False
    Else
        mobjValid.BreakRule "PHONE", False
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
    colClassErrors.Add "Initials too short", "INITIALS"
    colClassErrors.Add "Phone missing or too short", "PHONE"
    colClassErrors.Add "Account number too short", "ACNO"
    colClassErrors.Add "Note too short", "NOTE"
    colClassErrors.Add "Missing duplicate checking field", "CONTROL"
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

