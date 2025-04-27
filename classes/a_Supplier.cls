VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(strMsg As String)
Event DefaultAddressChanged()
Private moDefaultCurrency As a_Currency
Private mudtProps As SupplierProps
Private mcolStack As Collection
Private WithEvents mcolAddresses As ch_Addresses
Attribute mcolAddresses.VB_VarHelpID = -1
Private oApproAddress As a_Address
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oOrderToAddress As a_Address
Private mchDeals As ch_Deal
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mApproAddressIdx As String
Private mBilltoAddressIdx As String
Private mDelToAddressIdx As String
Private mOrderToAddressIdx As String
Private tlDispatchModes As z_TextList


Private Sub mcolAddresses_FirstAddressAdded()
    SetBillToAddressidx Me.Addresses(1).Key
    SetApproAddressidx Me.Addresses(1).Key
    SetDelToAddressidx Me.Addresses(1).Key
    SetOrderToAddressidx Me.Addresses(1).Key
    RaiseEvent DefaultAddressChanged
End Sub

Private Function GetState() As String
  Dim udtData As SupplierData
  LSet udtData = mudtProps
  GetState = udtData.buffer
End Function

Private Sub SetState(buffer As String)
Dim udtData As SupplierData
    udtData.buffer = buffer
    LSet mudtProps = udtData
    mobjValid.GetStatus
End Sub
Public Function GetSuperState() As String
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    mudtProps.DefaultCurrencyID = moDefaultCurrency.ID
    .WriteProperty "State", GetState
    .WriteProperty "Addresses", mcolAddresses.GetSuperState
    .WriteProperty "Deals", mchDeals.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Public Sub SetSuperState(buffer As String)
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
    mcolAddresses.SetSuperState .ReadProperty("Addresses")
    mchDeals.SetSuperState .ReadProperty("Deals")
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
    For Each moDefaultCurrency In oPC.Configuration.Currencies
        If moDefaultCurrency.ID = mudtProps.DefaultCurrencyID Then
            Exit For
        End If
        If moDefaultCurrency Is Nothing Then Set moDefaultCurrency = oPC.Configuration.DefaultCurrency
    Next
  Set objPB = Nothing
End Sub

Public Sub BeginEdit()
  mcolAddresses.BeginEdit
  mchDeals.BeginEdit
  mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mcolAddresses.CancelEdit
  mchDeals.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Public Sub ApplyEdit(lngResult As Long)
    On Error GoTo errHandler
Dim strTemp As String
Dim OpenResult As Integer
  Dim objPersist As a_Supplier_P
  lngResult = 0
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
 ' oPC.COShort.BeginTrans
  If mcolStack.Count = 0 Then Err.Raise 445
  If Not oBillToAddress Is Nothing Then Me.SetPhone oBillToAddress.Phone
  mcolAddresses.ApplyEdit
  mchDeals.ApplyEdit
  Set objPersist = New a_Supplier_P
  If mudtProps.IsDeleted Then
    objPersist.DeleteSupplier mudtProps.SupplierID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    strTemp = objPersist.Save(GetSuperState)
    If strTemp = "DUPLICATE" Then
     '   oPC.COShort.RollbackTrans
        mcolAddresses.BeginEdit
        mchDeals.BeginEdit
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
 ' oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Sub
   ' Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.ApplyEdit(lngResult)", lngResult, , , "Line number,BrokenRules", Array(Erl(), mobjValid.AllBrokenRules)
End Sub

Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mcolAddresses = New ch_Addresses
    Set mchDeals = New ch_Deal
    mudtProps.IsNew = True
    mudtProps.VATable = True
    Set colClassors = New Collection
    LoadClassorsCollection
    mBilltoAddressIdx = 1
    Set moDefaultCurrency = oPC.Configuration.DefaultCurrency
    Set tlDispatchModes = New z_TextList
    tlDispatchModes.Load ltDispatchModes
    mobjValid.RuleBroken "NAME", True

End Sub

Private Sub Class_Terminate()
    If mcolStack.Count > 0 Then Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set mcolAddresses = Nothing
    Set mchDeals = Nothing
    Set mcolStack = Nothing
End Sub


Private Sub mobjValid_BrokenRule(oRS As String)
    RaiseEvent Valid(TranslateErrors(oRS))
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    RaiseEvent Valid(TranslateErrors(oRS))
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid("")
End Sub
Private Sub mobjValid_Status(pMsg As String)
    RaiseEvent Valid(TranslateErrors(pMsg))
End Sub
Public Function Load(Optional SupplierID As Long = 0, Optional pACNO As String = "", Optional pDefaultPhone As String = "") As Boolean
Dim objPersist As a_Supplier_P
Dim tmpBuffer As String
Dim i As Integer
    If mcolStack.Count > 0 Then Err.Raise 445
    mudtProps.IsNew = False
    mBilltoAddressIdx = 1 'just in case
    If mudtProps.DefaultETA = 0 Then
        mudtProps.DefaultETA = 30
    End If
    Set objPersist = New a_Supplier_P
    Load = True
    tmpBuffer = ""
    If SupplierID > 0 Then
        tmpBuffer = objPersist.Fetch(SupplierID)
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
        Set objPersist = Nothing
        mobjValid.RuleBroken "Name", False
        mobjValid.RuleBroken "Phone", False
        mobjValid.RuleBroken "TP", False
    End If
    
    Set tlDispatchModes = New z_TextList
    tlDispatchModes.Load ltDispatchModes
    
End Function
Public Property Get DispatchModes() As z_TextList
    Set DispatchModes = tlDispatchModes
End Property

Public Property Get SupplierID() As Long
  SupplierID = mudtProps.SupplierID
End Property
Public Sub SetApproAddressidx(val As String)
Dim oAdd As a_Address
    mApproAddressIdx = val
    Set oApproAddress = Nothing
    Set oApproAddress = Me.Addresses(mApproAddressIdx)
    For Each oAdd In Addresses
        oAdd.Appro = False
    Next
    Addresses(mApproAddressIdx).Appro = True
End Sub
Public Property Get ApproAddressIdx() As Integer
    ApproAddressIdx = mApproAddressIdx
End Property
Public Property Get ApproAddress() As a_Address
    Set ApproAddress = oApproAddress
End Property

Public Sub SetBillToAddressidx(val As String)
Dim oAdd As a_Address
    mBilltoAddressIdx = val
    Set oBillToAddress = Nothing
    Set oBillToAddress = Me.Addresses(mBilltoAddressIdx)
    For Each oAdd In Addresses
        oAdd.BillTo = False
    Next
    Addresses(mBilltoAddressIdx).BillTo = True

    SetDirty True
End Sub
Public Property Get BillToAddressIdx() As Integer
    BillToAddressIdx = mBilltoAddressIdx
End Property
Public Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property

Public Sub SetDelToAddressidx(val As String)
Dim oAdd As a_Address
    mDelToAddressIdx = val
    Set oDelToAddress = Nothing
    Set oDelToAddress = Me.Addresses(mDelToAddressIdx)
    For Each oAdd In Addresses
        oAdd.DelTo = False
    Next
    Addresses(mDelToAddressIdx).DelTo = True
End Sub
Public Property Get DelTOAddressIdx() As Integer
    DelTOAddressIdx = mDelToAddressIdx
End Property
Public Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property

Public Sub SetOrderToAddressidx(val As String)
Dim oAdd As a_Address
    mOrderToAddressIdx = val
    Set oOrderToAddress = Nothing
    Set oOrderToAddress = Me.Addresses(mOrderToAddressIdx)
    For Each oAdd In Addresses
        oAdd.OrderTo = False
    Next
    Addresses(mOrderToAddressIdx).OrderTo = True
End Sub
Public Property Get OrderToAddressIdx() As Integer
    OrderToAddressIdx = mOrderToAddressIdx
End Property
Public Property Get OrderToAddress() As a_Address
    Set OrderToAddress = oOrderToAddress
End Property
Public Property Get ID() As Long
    ID = mudtProps.SupplierID
End Property
Public Property Let ParentSupplierID(val As Long)
    mudtProps.ParentSupplierID = val
End Property
Public Property Get ParentSupplierID() As Long
    ParentSupplierID = mudtProps.ParentSupplierID
End Property
Public Property Get ParentSupplierName() As String
    ParentSupplierName = mudtProps.ParentSupplierName
End Property
Public Property Get ReturnStartMonths() As Integer
    ReturnStartMonths = mudtProps.ReturnStartMonths
End Property
Public Function SetReturnStartMonths(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    bOK = SetField_INTEGER(mudtProps.ReturnStartMonths, val, "RETURNPERIOD", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("RETURNPERIOD")
    End If
    SetReturnStartMonths = bOK
End Function
Public Property Get ReturnEndMonths() As Integer
    ReturnEndMonths = mudtProps.ReturnEndMonths
End Property
Public Function SetReturnEndMonths(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    bOK = SetField_INTEGER(mudtProps.ReturnEndMonths, val, "RETURNPERIOD", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("RETURNPERIOD")
    End If
    SetReturnEndMonths = bOK
End Function

Public Property Get Addresses() As ch_Addresses
    Set Addresses = mcolAddresses
End Property
Public Property Get Deals() As ch_Deal
    Set Deals = mchDeals
End Property
Public Property Get Phone() As String
  Phone = FNS(mudtProps.Phone)
End Property
Public Function SetPhone(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.Phone, val, "PHONE", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("PHONE")
    End If
    SetPhone = bOK
End Function
Public Property Get Note() As String
  Note = FNS(mudtProps.Note)
End Property
Public Function SetNote(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.Note, val, "NOTE", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("NOTE")
    End If
    SetNote = bOK
End Function
Public Property Get NameAndCode(pLen As Integer) As String
Dim strNameandCode As String
    strNameandCode = Left(Trim(mudtProps.Name), pLen) & IIf(Len(Trim(mudtProps.Name)) > pLen, "...  ", "  ")
    If AcNo > "" Then
        strNameandCode = strNameandCode & "(" & AcNo & ")"
    End If
    NameAndCode = strNameandCode
End Property
Public Property Get Name() As String
  Name = FNS(mudtProps.Name)
End Property
Public Function SetName(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.Name, val, "NAME", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("NAME")
    End If
    SetName = bOK
End Function
Public Property Get GFXNumber() As String
  GFXNumber = Format(FNS(mudtProps.GFXNumber), "#############")
End Property
Public Function SetGFXNumber(val As String) As Boolean
Dim bOK As Boolean
    'pad with zeroes to right
    val = val & String(13 - Len(val), 48)
    bOK = SetField_STRING(mudtProps.GFXNumber, val, "GFXNUMBER", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("GFXNUMBER")
    End If
    SetGFXNumber = bOK
End Function
Public Property Get EDIType() As String
  EDIType = FNS(mudtProps.EDIType)
End Property
Public Function SetEDIType(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.EDIType, val, "EDIType", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("EDIType")
    End If
    SetEDIType = bOK
End Function


Public Property Get PO_FTPAddress() As String
  PO_FTPAddress = FNS(mudtProps.PO_FTPAddress)
End Property
Public Function SetPO_FTPAddress(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.PO_FTPAddress, val, "PO_FTPAddress", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("PO_FTPAddress")
    End If
    SetPO_FTPAddress = bOK
End Function

Public Property Get PO_FTPUser() As String
  PO_FTPUser = FNS(mudtProps.PO_FTPUser)
End Property
Public Function SetPO_FTPUser(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.PO_FTPUser, val, "PO_FTPUser", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("PO_FTPUser")
    End If
    SetPO_FTPUser = bOK
End Function

Public Property Get PO_FTPPassword() As String
  PO_FTPPassword = FNS(mudtProps.PO_FTPPassword)
End Property
Public Function SetPO_FTPPassword(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.PO_FTPPassword, val, "PO_FTPPassword", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("PO_FTPPassword")
    End If
    SetPO_FTPPassword = bOK
End Function

Public Property Get INV_FTPAddress() As String
  INV_FTPAddress = FNS(mudtProps.INV_FTPAddress)
End Property
Public Function SetINV_FTPAddress(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.INV_FTPAddress, val, "INV_FTPAddress", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("INV_FTPAddress")
    End If
    SetINV_FTPAddress = bOK
End Function
Public Property Get INV_FTPUser() As String
  INV_FTPUser = FNS(mudtProps.INV_FTPUser)
End Property
Public Function SetINV_FTPUser(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.INV_FTPUser, val, "INV_FTPUser", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("INV_FTPUser")
    End If
    SetINV_FTPUser = bOK
End Function
Public Property Get INV_FTPPassword() As String
  INV_FTPPassword = FNS(mudtProps.INV_FTPPassword)
End Property
Public Function SetINV_FTPPassword(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.INV_FTPPassword, val, "INV_FTPPassword", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("INV_FTPPassword")
    End If
    SetINV_FTPPassword = bOK
End Function

Public Property Get INV_FTPFolder() As String
  INV_FTPFolder = FNS(mudtProps.INV_FTPFolder)
End Property
Public Function SetINV_FTPFolder(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.INV_FTPFolder, val, "INV_FTPFolder", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("INV_FTPFolder")
    End If
    SetINV_FTPFolder = bOK
End Function

Public Property Get PO_FTPFolder() As String
  PO_FTPFolder = FNS(mudtProps.PO_FTPFolder)
End Property
Public Function SetPO_FTPFolder(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.PO_FTPFolder, val, "PO_FTPFolder", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("PO_FTPFolder")
    End If
    SetPO_FTPFolder = bOK
End Function

Public Property Get ConversionToLocalFactor() As Double
    On Error GoTo errHandler
    ConversionToLocalFactor = mudtProps.ConversionToLocalFactor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.ConversionToLocalFactor"
End Property
Public Property Get ConversionToLocalFactorF() As String
    On Error GoTo errHandler
    ConversionToLocalFactorF = Format(mudtProps.ConversionToLocalFactor, "##0.000")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.ConversionToLocalFactorF"
End Property
Public Function SetConversionToLocalFactor(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.ConversionToLocalFactor
    bOK = SetField_DOUBLE(mudtProps.ConversionToLocalFactor, val, "SetConversionToLocalFactor", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.ConversionToLocalFactor) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("SetConversionToLocalFactor")
    End If
    SetConversionToLocalFactor = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.SetConversionToLocalFactor(val)", val
End Function



Public Property Get UseStatus() As enUseStatus
  UseStatus = val(FNS(mudtProps.UseStatus))
End Property
Public Property Let UseStatus(val As enUseStatus)
  mudtProps.UseStatus = val
End Property

Public Property Get DefaultETA() As Integer
Dim iETA As Integer

  DefaultETA = FNN(mudtProps.DefaultETA)
      
End Property
Public Property Let DefaultETA(val As Integer)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultETA = val
    SetDirty True
End Property
Public Function SetDefaultETA(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    bOK = SetField_INTEGER(mudtProps.DefaultETA, val, "DEFAULTETA", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("DEFAULTETA")
    End If
    SetDefaultETA = bOK
End Function
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

Public Property Get DispatchMode() As String
    DispatchMode = FNS(mudtProps.DispatchMode)
End Property
Public Property Get DispatchModeID() As Long
    DispatchModeID = FNS(mudtProps.DispatchModeID)
End Property
Public Function SetDispatchModeID(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
        mudtProps.DispatchModeID = tlDispatchModes.Key(val)
End Function


Public Property Get DefaultCurrency() As a_Currency
    If moDefaultCurrency Is Nothing Then
        Set moDefaultCurrency = oPC.Configuration.DefaultCurrency
    End If
    Set DefaultCurrency = moDefaultCurrency
End Property
Public Function SetDefaultCurrency(val As a_Currency) As Boolean
    Set moDefaultCurrency = val
End Function
Public Property Get AcNo() As String
  AcNo = FNS(mudtProps.AcNo)
End Property
Public Function SetAcNO(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_STRING(mudtProps.AcNo, UCase(val), "ACNO", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("ACNO")
    End If
    SetAcNO = bOK
End Function
Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property
Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = (mcolStack.Count > 0)
End Property
Public Property Get IsValid() As Boolean
    IsValid = (mobjValid.Count = 0)
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mudtProps.IsDirty
End Property
Public Sub DeleteSupplier()
    mudtProps.IsDeleted = True
    SetDirty True
End Sub


Public Property Get DateRecordAddedF() As String
    If mudtProps.DateRecordAdded = CDate(0) Then
        DateRecordAddedF = "unknown"
    Else
        DateRecordAddedF = Format(mudtProps.DateRecordAdded, "dd/mm/yyyy")
    End If
End Property
Public Property Get DateRecordLastChangedF() As String
    If mudtProps.DateLastModified = CDate(0) Then
        DateRecordLastChangedF = "unknown"
    Else
        DateRecordLastChangedF = Format(mudtProps.DateLastModified, "dd/mm/yyyy")
    End If
End Property
Public Property Let VATable(val As Boolean)
    mudtProps.VATable = val
End Property
Public Property Get VATable() As Boolean
    VATable = mudtProps.VATable
End Property
Public Property Let DoNotOrderFrom(val As Boolean)
    mudtProps.DoNotOrderFrom = val
End Property
Public Property Get DoNotOrderFrom() As Boolean
    DoNotOrderFrom = mudtProps.DoNotOrderFrom
End Property
Public Function Validate(pMsg As String) As Boolean
    pMsg = mobjValid.AllBrokenRules
    Validate = (pMsg = "")
End Function
Private Function ValidateObject(pFld As String) As Boolean
Dim bValid As Boolean
    bValid = True
    If Not (Len(Trim(mudtProps.Name)) > 2) Then
        If pFld = "NAME" Then bValid = False
        mobjValid.BreakRule "NAME", True
    Else
        mobjValid.BreakRule "NAME", False
    End If
    If FNS(mudtProps.EDIType) <> "SA" And FNS(mudtProps.EDIType) <> "E9" And FNS(mudtProps.EDIType) <> "" Then
        If pFld = "EDIType" Then bValid = False
        mobjValid.BreakRule "EDIType", True
    Else
        mobjValid.BreakRule "EDIType", False
    End If
    If Not (Len(Trim(mudtProps.AcNo)) > 2 Or Len(Trim(mudtProps.AcNo)) = 0) Then
        If pFld = "ACNO" Then bValid = False
        mobjValid.BreakRule "ACNO", True
    Else
        mobjValid.BreakRule "ACNO", False
    End If
    If Not (Len(Trim(mudtProps.Note)) > 2 Or Len(Trim(mudtProps.Note)) = 0) Then
        If pFld = "NOTE" Then bValid = False
        mobjValid.BreakRule "NOTE", True
    Else
        mobjValid.BreakRule "NOTE", False
    End If
    If mudtProps.ReturnStartMonths > mudtProps.ReturnEndMonths Then
        If pFld = "RETURNPERIOD" Then bValid = False
        mobjValid.BreakRule "RETURNPERIOD", True
    Else
        mobjValid.BreakRule "RETURNPERIOD", False
    End If
'    If mudtProps.ReturnEndMonths < 2 Then
'        If pFld = "RETURNEND" Then bValid = False
'        mobjValid.BreakRule "RETURNEND", True
'    Else
'        mobjValid.BreakRule "RETURNEND", False
'    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
End Function
Private Function TranslateErrors(pRawors As String) As String
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
End Function
Private Sub LoadClassorsCollection()
    Set colClassors = New Collection
    colClassors.Add "Name too short", "NAME"
    colClassors.Add "Phone missing or too short", "PHONE"
    colClassors.Add "Account number too short", "ACNO"
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "Address missing", "ADDRESS"
    colClassors.Add "Invalid return period", "RETURNPERIOD"
    colClassors.Add "Invalid EDI TYPE", "EDITYPE"
End Sub
Private Sub SetDirty(pVal As Boolean)
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
End Sub

Public Function SupplierIndexClashes() As Boolean
Dim oSM As New a_Supplier_P
    
    If Me.IsNew Then
        SupplierIndexClashes = oSM.SupplierIndexClashes(0, Me.AcNo)
    Else
        SupplierIndexClashes = oSM.SupplierIndexClashes(Me.ID, Me.AcNo)
    End If

End Function
'-----------------------------------
Public Property Get TermsType() As String
  TermsType = FNS(mudtProps.TermsType)
End Property
Public Property Get TermsTypeF() As String
    Select Case FNS(mudtProps.TermsType)
    Case "doi"
        TermsTypeF = "Date of invoice"
    Case "eom"
        TermsTypeF = "End of month"
    End Select
End Property
Public Function SetTermsType(val As String) As Boolean
Dim bOK As Boolean
    Select Case val
    Case "1"
        val = "eom"
    Case "2"
        val = "doi"
    Case Else
        val = ""
    End Select
    bOK = SetField_STRING(mudtProps.TermsType, val, "TermsType", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("TermsType")
    End If
    SetTermsType = bOK
End Function

Public Property Get SettlementTermsTypeF() As String
    Select Case FNS(mudtProps.SettlementTermsType)
    Case "rog"
        SettlementTermsTypeF = "Receipt of goods"
    Case "doi"
        SettlementTermsTypeF = "Date of invoice"
    End Select
End Property
Public Function SetSettlementTermsType(val As String) As Boolean
Dim bOK As Boolean
    Select Case val
    Case "1"
        val = "doi"
    Case "2"
        val = "rog"
    Case Else
        val = ""
    End Select
    bOK = SetField_STRING(mudtProps.SettlementTermsType, val, "SettlementTermsType", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("SettlementTermsType")
    End If
    SetSettlementTermsType = bOK
End Function

Public Property Get Terms() As Integer
  Terms = FNN(mudtProps.Terms)
End Property
Public Property Let Terms(val As Integer)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Terms = val
    SetDirty True
End Property
Public Function SetTerms(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    bOK = SetField_INTEGER(mudtProps.Terms, val, "Terms", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("Terms")
    End If
    SetTerms = bOK
End Function
'claimNeedsApproval
Public Property Get SettlementTerms() As Integer
  SettlementTerms = FNN(mudtProps.SettlementTerms)
End Property
Public Property Let SettlementTerms(val As Integer)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SettlementTerms = val
    SetDirty True
End Property
Public Function SetSettlementTerms(val As String) As Boolean
Dim bOK As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    bOK = SetField_INTEGER(mudtProps.SettlementTerms, val, "SettlementTerms", mcolStack.Count)
    If bOK Then
        bOK = ValidateObject("SettlementTerms")
    End If
    SetSettlementTerms = bOK
End Function

Public Property Get SettlementDiscount() As Double
    On Error GoTo errHandler
    SettlementDiscount = mudtProps.SettlementDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.SettlementDiscount"
End Property
Public Property Get SettlementDiscountF() As String
    On Error GoTo errHandler
    SettlementDiscountF = Format(mudtProps.SettlementDiscount, "##0.000")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.SettlementDiscountF"
End Property
Public Function SetSettlementDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.SettlementDiscount
    bOK = SetField_DOUBLE(mudtProps.SettlementDiscount, val, "SettlementDiscount", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.SettlementDiscount) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("SettlementDiscount")
    End If
    SetSettlementDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier.SetSettlementDiscount(val)", val
End Function

Public Property Get TermsF() As String
    Select Case mudtProps.Terms
    Case 0
        TermsF = "on receipt"
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

Public Property Get ClaimNeedsApproval() As Boolean
  ClaimNeedsApproval = FNB(mudtProps.ClaimNeedsApproval)
End Property
Public Property Let ClaimNeedsApproval(val As Boolean)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ClaimNeedsApproval = val
    SetDirty True
End Property

