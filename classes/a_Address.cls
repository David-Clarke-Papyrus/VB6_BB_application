VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Address"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Note
Event Valid(pors As String, IsValid As Boolean)
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private tlCountries As z_TextList
Private mudtProps As AddressProps
Private mcolStack As Collection
Private mobjParent As ch_Addresses
Private mAddresses As ch_Addresses
Private colClassors As Collection
Private mKey As String
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property


Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As AddressData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As AddressData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
  mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.SetSuperState(buffer)", buffer
End Sub

Friend Sub Initialize(Parent As ch_Addresses)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
  SetParent Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Initialize(Parent)", Parent
End Sub

Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.ID"
End Property

Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.TPID"
End Property
Public Property Let TPID(val As Long)
    On Error GoTo errHandler
    mudtProps.TPID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.TPID(val)", val
End Property

Public Property Get CountryID() As Long
    On Error GoTo errHandler
    CountryID = mudtProps.CountryID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.CountryID"
End Property
Public Property Let CountryID(val As Long)
    On Error GoTo errHandler
    mudtProps.CountryID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.CountryID(val)", val
End Property

Public Property Get Description() As String
    On Error GoTo errHandler
    Description = FNS(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Description"
End Property
Public Property Let Description(val As String)
    On Error GoTo errHandler
    mudtProps.Description = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Description(val)", val
End Property
Public Function SetDescription(val As String) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Description) Then
        Err.Raise 384
    End If
    mudtProps.Description = val
    SetDescription = ValidateObject(val, "DESCRIPTION")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    ErrPreserve
    If Err.Number <> 383 And Err.Number <> 384 Then
        If ErrMustStop Then Debug.Assert False: Resume
        ErrorIn "a_Address.SetDescription(val)", val
    End If
End Function

Public Property Get Line1() As String
    On Error GoTo errHandler
    Line1 = FNS(mudtProps.Line1)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Line1"
End Property
Public Function SetLine1(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Line1) Then
        Err.Raise 384
    End If
    mudtProps.Line1 = val
    SetLine1 = ValidateObject(val, "LINE1")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Line2() As String
    On Error GoTo errHandler
    Line2 = FNS(mudtProps.Line2)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Line2"
End Property
Public Function SetLine2(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Line2) Then
        Err.Raise 384
    End If
    SetLine2 = True
    mudtProps.Line2 = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Line3() As String
    On Error GoTo errHandler
    Line3 = FNS(mudtProps.Line3)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Line3"
End Property
Public Function SetLine3(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Line3) Then
        Err.Raise 384
    End If
    SetLine3 = True
    mudtProps.Line3 = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Line4() As String
    On Error GoTo errHandler
    Line4 = FNS(mudtProps.Line4)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Line4"
End Property
Public Function SetLine4(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Line4) Then
        Err.Raise 384
    End If
    SetLine4 = True
    mudtProps.Line4 = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Line5() As String
    On Error GoTo errHandler
    Line5 = FNS(mudtProps.Line5)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Line5"
End Property
Public Function SetLine5(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Line5) Then
        Err.Raise 384
    End If
    SetLine5 = True
    mudtProps.Line5 = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Line6() As String
    Line6 = FNS(mudtProps.Line6)
    Exit Property
End Property
Public Function SetLine6(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Line6) Then
        Err.Raise 384
    End If
    SetLine6 = True
    mudtProps.Line6 = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get pCode() As String
    On Error GoTo errHandler
    pCode = FNS(mudtProps.pCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.pCODE"
End Property
Public Function SetPCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.pCode) Then
        Err.Raise 384
    End If
    mudtProps.pCode = val
    SetPCode = ValidateObject(val, "PCODE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Sub SetAddress(val As String)
Dim ar() As String
Dim i As Integer
Dim IsStrArrayEmpty As Boolean

    ar = Split(val, Chr(13) & Chr(10))
    IsStrArrayEmpty = True
    For i = 0 To UBound(ar)
        If LenB(ar(i)) Then      ' if any member of ArrS is not blank then fail (LenB is faster than Len)
            IsStrArrayEmpty = False
            Exit For
        End If
    Next
    If IsStrArrayEmpty Then Exit Sub
    Me.SetAddressee ar(0)
    If UBound(ar) > 0 Then
        Me.SetLine1 ar(1)
    End If
    If UBound(ar) > 1 Then
        Me.SetLine2 ar(2)
    End If
    If UBound(ar) > 2 Then
        Me.SetLine3 ar(3)
    End If
    If UBound(ar) > 3 Then
        Me.SetLine4 ar(4)
    End If
    If UBound(ar) > 4 Then
        Me.SetLine5 ar(5)
    End If
    If UBound(ar) > 5 Then
        Me.SetLine6 ar(6)
    End If
End Sub

Public Property Get Addressee() As String
    On Error GoTo errHandler
    Addressee = FNS(mudtProps.Addressee)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Addressee"
End Property
Public Function SetAddressee(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Addressee) Then
        Err.Raise 384
    End If
    mudtProps.Addressee = val
    SetAddressee = ValidateObject(val, "ADDRESSEE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get Delivery() As String
    On Error GoTo errHandler
    Delivery = FNS(mudtProps.Delivery)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Delivery"
End Property
Public Function SetDelivery(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Delivery) Then
        Err.Raise 384
    End If
    mudtProps.Delivery = val
    SetDelivery = ValidateObject(val, "DELIVERY")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get BusPhone() As String
    On Error GoTo errHandler
    BusPhone = FNS(mudtProps.BusPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.BusPhone"
End Property
Public Property Get BusPhoneF() As String
    If Len(FNS(mudtProps.BusPhone)) = 10 Then
        BusPhoneF = Format(FNS(mudtProps.BusPhone), "\(@@@\)\ @@@\-@@@@")
    Else
        BusPhoneF = FNS(mudtProps.BusPhone)
    End If
End Property

Public Function SetBusPhone(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.BusPhone) Then
        Err.Raise 384
    End If
    mudtProps.BusPhone = val
    SetBusPhone = ValidateObject(val, "BUSPHONE")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Phone() As String
    On Error GoTo errHandler
    Phone = FNS(mudtProps.Phone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Phone"
End Property
Public Property Get Phonef() As String
    If Len(FNS(mudtProps.Phone)) = 10 Then
        Phonef = Format(FNS(mudtProps.Phone), "\(@@@\)\ @@@\-@@@@")
    Else
        Phonef = FNS(mudtProps.Phone)
    End If
End Property
Public Property Get PhoneandFax() As String
    On Error GoTo errHandler
Dim strPhone As String
Dim strFax As String
Dim strOut As String

    strPhone = FNS(mudtProps.Phone)
    strFax = FNS(mudtProps.Fax)
    strOut = ""
    If strPhone > "" Then
        strOut = "Phone: " & strPhone
    End If
    If strFax > "" Then
        If Len(strOut) > 0 Then strOut = strOut & "   "
        strOut = strOut & "Fax: " & strFax
    End If
    PhoneandFax = strOut
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.PhoneandFax"
End Property
Public Function SetPhone(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Phone) Then
        Err.Raise 384
    End If
    mudtProps.Phone = val
    If Not mAddresses.Customer Is Nothing Then
        If Me Is mAddresses.BillTOAddress Then
            mAddresses.Customer.SetPhone val
        End If
    End If
    SetPhone = ValidateObject(val, "PHONE")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Fax() As String
    On Error GoTo errHandler
    Fax = FNS(mudtProps.Fax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Fax"
End Property
Public Property Get FaxF() As String
    If Len(FNS(mudtProps.Fax)) = 10 Then
        FaxF = Format(FNS(mudtProps.Fax), "\(@@@\)\ @@@\-@@@@")
    Else
        FaxF = FNS(mudtProps.Fax)
    End If
End Property

Public Function SetFax(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Fax) Then
        Err.Raise 384
    End If
    mudtProps.Fax = val
    SetFax = ValidateObject(val, "FAX")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get EMail() As String
    On Error GoTo errHandler
    EMail = FNS(mudtProps.EMail)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Email"
End Property
Public Function SetEmail(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.EMail) Then
        Err.Raise 384
    End If
    mudtProps.EMail = val
    SetEmail = ValidateObject(val, "EMAIL")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get PostageType() As Integer
    On Error GoTo errHandler
    PostageType = FNN(mudtProps.PostageType)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.PostageType"
End Property
Public Property Let PostageType(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PostageType = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.PostageType(val)", val
End Property
Public Property Get Category() As Integer
    On Error GoTo errHandler
    Category = FNN(mudtProps.Category)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Category"
End Property
Public Property Let Category(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Category = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Category(val)", val
End Property

Public Property Get GetsCatalogue() As Boolean
    On Error GoTo errHandler
    GetsCatalogue = FNB(mudtProps.GetsCatalogue)
    If FNB(mudtProps.GetsCatalogue) Then mudtProps.ForMailing = True

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.GetsCatalogue"
End Property
Public Property Let GetsCatalogue(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.GetsCatalogue = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.GetsCatalogue(val)", val
End Property

Public Property Get ForMailing() As Boolean
    On Error GoTo errHandler
    ForMailing = mudtProps.ForMailing
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.ForMailing"
End Property
Public Property Let ForMailing(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ForMailing = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.ForMailing(val)", val
End Property
Public Function SetForMailing(val As Boolean) As Boolean
    On Error GoTo errHandler
If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ForMailing = val
    mudtProps.IsDirty = True
    SetForMailing = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.SetForMailing(val)", val
End Function

Public Property Get Appro() As Boolean
    On Error GoTo errHandler
    Appro = mudtProps.Appro
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Appro"
End Property
Public Property Let Appro(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Appro = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Appro(val)", val
End Property
Public Property Get BillTo() As Boolean
    On Error GoTo errHandler
    BillTo = mudtProps.BillTo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.BillTo"
End Property
Public Property Let BillTo(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BillTo = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.BillTo(val)", val
End Property
Public Property Get DelTo() As Boolean
    On Error GoTo errHandler
    DelTo = mudtProps.DelTo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.DelTo"
End Property
Public Property Let DelTo(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DelTo = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.DelTo(val)", val
End Property
Public Property Get OrderTo() As Boolean
    On Error GoTo errHandler
    OrderTo = mudtProps.OrderTo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.OrderTo"
End Property
Public Property Let OrderTo(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OrderTo = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.OrderTo(val)", val
End Property

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.BeginEdit"
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
    ErrorIn "a_Address.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddAddress Me
    Set mAddresses = mobjParent
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.ApplyEdit"
End Sub
Public Sub SetParent(pCol As ch_Addresses)
    Set mAddresses = pCol
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set tlCountries = New z_TextList
    tlCountries.Load ltCountry
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mobjValid.BreakRule "LINE1", True
    mudtProps.IsNew = True
    Set colClassors = New Collection
  '  mudtProps.Addressee = "<Double-click to enter new details>"
    LoadClassorsCollection
    ValidateObject "", ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Class_Initialize"
End Sub
Public Sub GetStatus()
    On Error GoTo errHandler
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.GetSTatus"
End Sub
Public Property Get Countries() As z_TextList
    Set Countries = tlCountries
End Property
Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set tlCountries = Nothing
    Set colClassors = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(sRulesBroken As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(sRulesBroken), False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.mobjValid_BrokenRule(sRulesBroken)", sRulesBroken
End Sub
Private Sub mobjValid_NoBrokules()
    On Error GoTo errHandler
    RaiseEvent Valid("", True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.mobjValid_NoBrokules"
End Sub
Private Sub mobjValid_Status(sRulesBroken As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(sRulesBroken), (sRulesBroken = ""))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.mobjValid_Status(sRulesBroken)", sRulesBroken
End Sub
Public Sub Delete()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    mudtProps.IsDeleted = True
    mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Delete"
End Sub


Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.IsNew"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.IsEditing"
End Property
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.IsValid"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.IsDirty"
End Property
Public Function Validate(pMsg As String) As Boolean
    On Error GoTo errHandler
    pMsg = mobjValid.AllBrokenRules
    Validate = (pMsg = "")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.Validate(pMsg)", pMsg
End Function

Public Function AddressMailing() As String
    On Error GoTo errHandler
    AddressMailing = FNS(mudtProps.Addressee)  ' & IIf(bWithCode, " (" & mudtProps.cu & ")", "")
    If HasNonEmptyString(mudtProps.Line1) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.Line1)
 '   Else
 '       GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line2) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.Line2)
 '   Else
 '       GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line3) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.Line3)
 '   Else
 '       GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line4) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.Line4)
 '   Else
 '       GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line5) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.Line5)
 '   Else
 '       GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line6) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.Line6)
 '   Else
 '       GoTo CheckPostCodeAndCountry
    End If

CheckPostCodeAndCountry:
    If HasNonEmptyString(mudtProps.pCode) Then
        AddressMailing = AddressMailing & vbCrLf & FNS(mudtProps.pCode)
    End If
    If HasNonEmptyString(CountryName) And Me.CountryID <> oPC.Configuration.LocalCountryID Then
        AddressMailing = AddressMailing & vbCrLf & CountryName
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.AddressMailing"
End Function
Public Function AddressShort() As String
    On Error GoTo errHandler
    AddressShort = Trim(mudtProps.Addressee)
    If HasNonEmptyString(mudtProps.Line1) Then
        AddressShort = AddressShort & vbCrLf & Trim(mudtProps.Line1)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line2) Then
        AddressShort = AddressShort & vbCrLf & Trim(mudtProps.Line2)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line3) Then
        AddressShort = AddressShort & vbCrLf & Trim(mudtProps.Line3)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line4) Then
        AddressShort = AddressShort & vbCrLf & Trim(mudtProps.Line4)
    End If
    If HasNonEmptyString(mudtProps.Line5) Then
        AddressShort = AddressShort & vbCrLf & Trim(mudtProps.Line5)
    End If
CheckPostCodeAndCountry:
    If HasNonEmptyString(CountryName) And Me.CountryID <> oPC.Configuration.LocalCountryID Then
        AddressShort = AddressShort & vbCrLf & CountryName
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.AddressShort"
End Function
Public Function AddressDelimited() As String
    On Error GoTo errHandler
    AddressDelimited = FNS(mudtProps.Addressee)
    If HasNonEmptyString(mudtProps.Line1) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.Line1)
    Else
        GoTo CheckPostCodeAndCountry
    End If
'    AddressDelimited = Trim(mudtProps.Line1)
    If HasNonEmptyString(mudtProps.Line2) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.Line2)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line3) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.Line3)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line4) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.Line4)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line5) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.Line5)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If HasNonEmptyString(mudtProps.Line6) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.Line6)
    Else
        GoTo CheckPostCodeAndCountry
    End If

CheckPostCodeAndCountry:
    If HasNonEmptyString(CountryName) And Me.CountryID <> oPC.Configuration.LocalCountryID Then
        AddressDelimited = AddressDelimited & "§" & CountryName
    End If
    If HasNonEmptyString(mudtProps.pCode) Then
        AddressDelimited = AddressDelimited & "§" & FNS(mudtProps.pCode)
    End If
    AddressDelimited = Trim(AddressDelimited)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.AddressDelimited"
End Function

Public Property Get CountryName() As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
    If Me.CountryID > 0 Then
        CountryName = Me.Countries.Item(CStr(Me.CountryID))
    Else
        CountryName = ""
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.CountryName"
End Property
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If Not (Len(Trim(mudtProps.Fax)) > 3 Or Len(Trim(mudtProps.Fax)) = 0) Then
        mobjValid.BreakRule "FAX", True
        If pFld = "FAX" Then bValid = False
    Else
        mobjValid.BreakRule "FAX", False
    End If
'    If Not (Len(FNS(mudtProps.Description)) > 2) Then
'        mobjValid.BreakRule "DESCRIPTION", True
'        If pFld = "DESCRIPTION" Then bValid = False
'    Else
'        mobjValid.BreakRule "DESCRIPTION", False
'    End If
    If Not (Len(FNS(mudtProps.Line1)) > 2 Or Len(FNS(mudtProps.Line1)) = 0) Then
        mobjValid.BreakRule "LINE1", True
        If pFld = "LINE1" Then bValid = False
    Else
        mobjValid.BreakRule "LINE1", False
    End If
    If Not (Len(FNS(mudtProps.pCode)) > 3 Or Len(FNS(mudtProps.pCode)) = 0) Then
        mobjValid.BreakRule "PCODE", True
        If pFld = "PCODE" Then bValid = False
    Else
        mobjValid.BreakRule "PCODE", False
    End If
    If Not (Len(FNS(mudtProps.Phone)) > 5 Or Len(FNS(mudtProps.Phone)) = 0) Then
        mobjValid.BreakRule "PHONE", True
        If pFld = "PHONE" Then bValid = False
    Else
        mobjValid.BreakRule "PHONE", False
    End If
    If Not (Len(FNS(mudtProps.Phone)) > 5 Or Len(FNS(mudtProps.Phone)) = 0) Then
        mobjValid.BreakRule "PHONE", True
        If pFld = "PHONE" Then bValid = False
    Else
        mobjValid.BreakRule "PHONE", False
    End If
    If Not (Len(FNS(mudtProps.BusPhone)) > 5 Or Len(FNS(mudtProps.BusPhone)) = 0) Then
        mobjValid.BreakRule "BUSPHONE", True
        If pFld = "BUSPHONE" Then bValid = False
    Else
        mobjValid.BreakRule "BUSPHONE", False
    End If
    If Not (Len(FNS(mudtProps.Fax)) > 5 Or Len(FNS(mudtProps.Fax)) = 0) Then
        mobjValid.BreakRule "FAX", True
        If pFld = "FAX" Then bValid = False
    Else
        mobjValid.BreakRule "FAX", False
    End If
    If Not (Len(FNS(mudtProps.EMail)) > 5 Or Len(FNS(mudtProps.EMail)) = 0) Then
        mobjValid.BreakRule "EMAIL", True
        If pFld = "EMAIL" Then bValid = False
    Else
        mobjValid.BreakRule "EMAIL", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.ValidateObject(pVal,pFld)", Array(pVal, pFld)
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
    ErrorIn "a_Address.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Address description too short", "DESCRIPTION"
    colClassors.Add "Address line one too short", "LINE1"
    colClassors.Add "Addressee too short", "ADDRESSEE"
    colClassors.Add "Phone too short", "PHONE"
    colClassors.Add "Business phone too short", "BUSPHONE"
    colClassors.Add "Fax too short", "FAX"
    colClassors.Add "Email too short", "EMAIL"
    colClassors.Add "PCode too short", "PCODE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.LoadClassorsCollection"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Address.SetDirty(pVal)", pVal
End Sub

