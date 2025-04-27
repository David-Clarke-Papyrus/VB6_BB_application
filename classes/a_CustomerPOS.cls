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
Event Dirty(bYesNO As Boolean)
Private mudtProps As CustomerPOSProps
Private mcolStack As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassErrors As Collection
'Dim XA As XArrayDB
'Dim XB As XArrayDB

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As CustomerPOSData
  LSet udtData = mudtProps
  GetState = udtData.Buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GetState"
End Function

Private Sub SetState(Buffer As String)
    On Error GoTo errHandler
Dim udtData As CustomerPOSData
    udtData.Buffer = Buffer
    LSet mudtProps = udtData
    mobjValid.GetSTatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetState(buffer)", Buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.GetSuperState"
End Function

Public Sub SetSuperState(Buffer As String)
    On Error GoTo errHandler
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim i As Integer

    If Buffer = "" Then Exit Sub
    Set objPB = New PropertyBag
    arBuffer = Buffer
    With objPB
      .Contents = arBuffer
      SetState .ReadProperty("State")
    End With
    
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetSuperState(buffer)", Buffer
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
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
Dim objPersist As a_CustomerPOS_P

  lngResult = 0
  If mcolStack.Count = 0 Then Err.Raise 445
  Set objPersist = New a_CustomerPOS_P
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject mudtProps.ID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    strTemp = objPersist.Save(GetSuperState)
    If strTemp = "DUPLICATE" Then
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
EXIT_Handler:
    Exit Sub
   ' Resume
ERR_Handler:
    lngResult = 101
    MsgBox "Unknown Error" & Error
    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.ApplyEdit(lngResult)", lngResult
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
  '  mudtProps.CustomerTypeID = "CAS"
    mudtProps.Name = ""
    mudtProps.IsNew = True
    Set colClassErrors = New Collection
    LoadClassErrorsCollection
    VATable = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then Err.Raise vbObjectError + 1001, , "State stack is not empty"
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

Public Function Load(Optional CustomerID As Long = 0, Optional pACNO As String = "", Optional pDefaultPhone As String = "") As Boolean
    On Error GoTo errHandler
Dim objPersist As a_CustomerPOS_P
Dim tmpBuffer As String
Dim i As Integer
    If mcolStack.Count > 0 Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_CustomerPOS_P
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
    CustomerID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.CustomerID"
End Property

Public Property Get CustomerType() As String
    CustomerType = Replace(UCase(FNS(mudtProps.CustomerTypeID)), "*", "")
End Property
Public Property Get CustomerTypeRaw() As String
    CustomerTypeRaw = UCase(FNS(mudtProps.CustomerTypeID))
End Property

Public Property Get CustomerTypeID() As String
    CustomerTypeID = FNS(CustomerType)
End Property

Public Property Get IsLoyaltyCustomer() As Boolean
    IsLoyaltyCustomer = FNB(((CustomerTypeID = oPC.GetLoyaltyCode()) And CustomerTypeID > ""))
End Property
Public Property Get CustomerDiscount() As Long
    CustomerDiscount = mudtProps.DefaultDiscount
End Property

Public Property Get CreditLimit() As Long
    CreditLimit = mudtProps.CreditLimit
End Property
Public Property Get CreditLimitF() As String
    CreditLimitF = Format(mudtProps.CreditLimit, oPC.CurrencyFormat)
End Property
Public Property Get Balance() As Long
    Balance = mudtProps.Balance
End Property
Public Property Get BalanceF() As String
    BalanceF = Format(mudtProps.Balance, oPC.CurrencyFormat)
End Property
Public Property Get Balances() As String
    Balances = FNS(mudtProps.Balances)
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
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = FNN(mudtProps.ID)
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

Public Property Get Initials() As String
    On Error GoTo errHandler
  Initials = Trim$(mudtProps.Initials)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Initials"
End Property
Public Function SetInitials(val As String) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetInitials = True
    If Len(val) > Len(mudtProps.Initials) Then
        Err.Raise 384
    End If
    mudtProps.Initials = val
    SetInitials = ValidateObject(val, "INITIALS")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetInitials(val)", val
End Function
Public Property Get Phone() As String
    On Error GoTo errHandler
  Phone = FNS(mudtProps.Phone)
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
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetPhone(val)", val
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
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetNote = True
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    mudtProps.Note = val
    SetNote = ValidateObject(val, "Note")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetNote(val)", val
End Function

Public Property Get title() As String
    On Error GoTo errHandler
  title = Trim$(mudtProps.title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.title"
End Property
Public Function SetTitle(val As String) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetTitle = True
    If Len(val) > Len(mudtProps.title) Then
        Err.Raise 384
    End If
    mudtProps.title = val
    SetTitle = ValidateObject(val, "TITLE")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetTitle(val)", val
End Function
Public Function SetCustomerTypeCasual() As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCustomerTypeCasual = True
    mudtProps.CustomerTypeID = "CAS"
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetCustomerTypeCasual"
End Function

Public Property Get Fullname() As String
    On Error GoTo errHandler
    Fullname = Name & IIf(Len(title) > 0 Or Len(Initials) > 0, ", ", "") & title & IIf(Len(title) > 0, " ", "") & Initials
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
Public Property Get NameAndCodeandType(ByVal pLen As Integer) As String
    On Error GoTo errHandler
Dim strNameandCode As String
Dim strDescription As String

    If Me.CustomerType = oPC.GetLoyaltyCode Then
        strDescription = "(Loyalty)"
    ElseIf Me.CustomerType = oPC.GetBookClubCode Then
        strDescription = "(Book club)"
    Else
        strDescription = ""
    End If
    pLen = pLen - Len(strDescription) - 1
    If pLen <= 0 Then pLen = 50
        strNameandCode = Left(Fullname, pLen) & IIf(Len(Fullname) > pLen, "...  ", "  ")
    If AcNo > "" Then
        strNameandCode = strNameandCode & "(" & AcNo & ")"
    End If
    NameAndCodeandType = strNameandCode & " " & strDescription
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.NameAndCodeandType(pLen)", pLen
End Property

Public Property Get Name() As String
    On Error GoTo errHandler
    Name = FNS(mudtProps.Name)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.Name"
End Property
Public Function SetName(val As String) As Boolean
    On Error GoTo errHandler
   ' If mcolStack.Count = 0 Then Err.Raise 383
    SetName = True
    If Len(val) > Len(mudtProps.Name) Then
        Err.Raise 384
    End If
    mudtProps.Name = val
    SetName = ValidateObject(val, "NAME")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetName(val)", val
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
    On Error GoTo errHandler
 '   If mcolStack.Count = 0 Then Err.Raise 383
    SetAcNO = True
    If Len(val) > Len(mudtProps.title) Then
        Err.Raise 384
    End If
    mudtProps.AcNo = UCase(val)
    SetAcNO = ValidateObject(val, "ACNO")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer.SetAcNO(val)", val
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
    If Not (Len(Trim(mudtProps.Name)) > 2 Or Len(Trim(mudtProps.Name)) = 0) Then
        mobjValid.BreakRule "NAME", True
        If pFld = "NAME" Then bValid = False
    Else
        mobjValid.BreakRule "NAME", False
    End If
    If Not (Len(Trim(mudtProps.title)) > 1 Or Len(Trim(mudtProps.title)) = 0) Then
        mobjValid.BreakRule "TITLE", True
        If pFld = "TITLE" Then bValid = False
    Else
        mobjValid.BreakRule "TITLE", False
    End If
    If Not (Len(Trim(mudtProps.AcNo)) > 2 Or Len(Trim(mudtProps.AcNo)) = 0) Then
        mobjValid.BreakRule "ACNO", True
        If pFld = "ACNO" Then bValid = False
    Else
        mobjValid.BreakRule "ACNO", False
    End If
    If Not (Len(Trim(mudtProps.Note)) > 2 Or Len(Trim(mudtProps.Note)) = 0) Then
        mobjValid.BreakRule "NOTE", True
        If pFld = "NOTE" Then bValid = False
    Else
        mobjValid.BreakRule "NOTE", False
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
Dim iMarker, istart As Integer
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
            istart = iMarker + 1
            iMarker = InStr(istart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = colClassErrors(MID(pRawErrors, istart, iMarker - istart))
            Else
                strRule = colClassErrors(MID(pRawErrors, istart))
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

