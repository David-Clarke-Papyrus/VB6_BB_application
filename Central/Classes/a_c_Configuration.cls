VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_c_Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mudtProps As ConfigProps
Private mcolStack As Collection
Private mcolCurrencies As ch_Currencies
Private mcolStores As ch_Store
Private mcolBICs As c_BICCodes
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private tlSections As z_TextList
Private tlProductTypes As z_TextList
Private tlWorkstations As z_TextList
Private tlPrinters As z_TextList

Dim tlStores As z_TextList
Private iBookfindStatus As Integer

Event Valid(errors As String, STATUS As Boolean)

Private colClassErrors As Collection
Private strErrors As String
Public Property Get Sections() As z_TextList
    Set Sections = tlSections
End Property
Public Property Get ProductTypes() As z_TextList
    Set ProductTypes = tlProductTypes
End Property
Public Property Get SupportsLoyaltyClub() As Boolean
    SupportsLoyaltyClub = FNB(mudtProps.SupportsLoyaltyClub)
End Property
Public Sub LoadStores_tl(pAllowAny As String)
    Set tlStores = Nothing
    Set tlStores = New z_TextList
    tlStores.Load ltStores, , pAllowAny
End Sub
Public Property Get Stores_tl() As z_TextList
    Set Stores_tl = tlStores
End Property

Public Sub RefreshProductTypes()
    On Error GoTo ErrHandler
    Set tlProductTypes = Nothing
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltProductType
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.RefreshProductTypes"
End Sub
Public Sub RefreshSections()
    On Error GoTo ErrHandler
    Set tlSections = Nothing
    Set tlSections = New z_TextList
    tlSections.Load ltDictionary, , dtCategory
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.RefreshSections"
End Sub

Private Function GetState() As String
    On Error GoTo ErrHandler
  Dim udtData As ConfigData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo ErrHandler
  Dim udtData As ConfigData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.SetState(buffer)", buffer
End Sub

Public Function GetSuperState() As String
    On Error GoTo ErrHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "Currencies", mcolCurrencies.GetSuperState
    .WriteProperty "Stores", mcolStores.GetSuperState
    .WriteProperty "BICs", mcolBICs.Load
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.GetSuperState"
End Function

Public Sub SetSuperState(buffer As String)
    On Error GoTo ErrHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mcolCurrencies.SetSuperState .ReadProperty("Currencies")
    mcolStores.SetSuperState .ReadProperty("Stores")
    mcolBICs.SetSuperState .ReadProperty("BICs")
  End With
  Set objPB = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.SetSuperState(buffer)", buffer
End Sub

Public Function Load() As Long
    On Error GoTo ErrHandler

    Dim objPersist As a_c_Configuration_P
    
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    ' code to load the object goes here
    Set objPersist = New a_c_Configuration_P
    SetSuperState objPersist.Fetch()
    Set objPersist = Nothing
    Set tlSections = New z_TextList
    Set tlProductTypes = New z_TextList
    tlSections.Load ltSectionsActive, , "<All>"
    tlProductTypes.Load ltProductTypeActive, , "<All>"
    
    LoadStores_tl ""
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.Load"
End Function
Public Function Stores() As ch_Store
    On Error GoTo ErrHandler
    Set Stores = mcolStores
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.Stores"
End Function
Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Set mcolStack = New Collection
    Set mcolCurrencies = New ch_Currencies
    Set mcolStores = New ch_Store
    Set mcolBICs = New c_BICCodes
    Set mobjValid = New z_BrokenRules
    Set colClassErrors = New Collection
    LoadClassErrorsCollection
  
    mudtProps.IsNew = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrHandler
    If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    
    Set mcolCurrencies = Nothing
    Set mcolStack = Nothing
    Set mcolStores = Nothing
    Set mcolBICs = Nothing
    Set mobjValid = Nothing
    Set colClassErrors = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.Class_Terminate"
End Sub
Private Sub LoadClassErrorsCollection()
    On Error GoTo ErrHandler
    Set colClassErrors = New Collection
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.LoadClassErrorsCollection"
End Sub
Public Function BICs() As c_BICCodes
    On Error GoTo ErrHandler
    Set BICs = mcolBICs
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.BICs"
End Function
Private Function TranslateErrors(pRawErrors As String) As String
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.TranslateErrors(pRawErrors)", pRawErrors
End Function

Private Sub mobjValid_BrokenRule(pErrors As String)
    On Error GoTo ErrHandler
    strErrors = TranslateErrors(pErrors)
    RaiseEvent Valid(TranslateErrors(pErrors), IsValid)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.mobjValid_BrokenRule(pErrors)", pErrors
End Sub
Private Sub mobjValid_RuleUnbroken(pErrors As String)
    On Error GoTo ErrHandler
    strErrors = TranslateErrors(pErrors)
    RaiseEvent Valid(TranslateErrors(pErrors), IsValid)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.mobjValid_RuleUnbroken(pErrors)", pErrors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo ErrHandler
    strErrors = ""
    RaiseEvent Valid("", True)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.mobjValid_NoBrokenRules"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo ErrHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.IsValid"
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo ErrHandler
  IsDirty = mudtProps.IsDirty Or mcolCurrencies.IsDirty
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.IsDirty"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo ErrHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.IsEditing"
End Property

Public Sub BeginEdit()
    On Error GoTo ErrHandler
    mcolStores.BeginEdit
  ' save object state
   mcolCurrencies.BeginEdit
   mcolStack.Add GetState
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.BeginEdit"
End Sub
Public Sub CancelEdit()
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    
  mcolCurrencies.CancelEdit
    mcolStores.CancelEdit
    mudtProps.IsDeleted = False
  ' restore object state
   With mcolStack
     SetState .Item(.Count)
     .Remove .Count
   End With
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.CancelEdit"
End Sub
Public Sub ApplyEdit(pErrors As String)
    On Error GoTo ErrHandler
Dim objPersist As a_c_Configuration_P
Dim strStatus As String

    If mcolStack.Count = 0 Then Err.Raise 445
    mcolCurrencies.ApplyEdit
    mcolStores.ApplyEdit
    Set objPersist = New a_c_Configuration_P
    If mudtProps.IsDeleted Then
      objPersist.DeleteObject 1
      mcolStack.Remove mcolStack.Count
      mudtProps.IsNew = True
      mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
      If Not IsValid Then Err.Raise 445
      mcolStack.Remove mcolStack.Count
      SetSuperState objPersist.Save(GetSuperState, strStatus)
      mudtProps.IsNew = False
    Else
      mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    mudtProps.IsDirty = False
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.ApplyEdit(pErrors)", pErrors
End Sub
Private Function GetSQLMeaning(pCODE As String) As String
    On Error GoTo ErrHandler
    Select Case pCODE
    Case "23000"
        GetSQLMeaning = "There is a disallowed duplicate value in a field"
    Case "22000"
        GetSQLMeaning = "A value is missing in a field"
    End Select
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.GetSQLMeaning(pCODE)", pCODE
End Function
Public Property Get LocalCountryID() As Long
    On Error GoTo ErrHandler
    LocalCountryID = mudtProps.LocalCountryID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.LocalCountryID"
End Property
Public Property Let LocalCountryID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LocalCountryID = val
    mudtProps.IsDirty = True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.LocalCountryID(val)", val
End Property


Public Property Get UsesBookfind() As Boolean
    On Error GoTo ErrHandler
    UsesBookfind = InStr(mudtProps.LookupSeq, "BF") > 0
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.UsesBookfind"
End Property


Public Property Get LookupSeq() As String
    On Error GoTo ErrHandler
    LookupSeq = FNS(mudtProps.LookupSeq)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.LookupSeq"
End Property
Public Function SetLookupSequence(val As String) As Boolean
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LookupSeq = UCase(val)
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.SetLookupSequence(val)", val
End Function
Public Property Get UsesWhitaker() As Boolean
    On Error GoTo ErrHandler
    UsesWhitaker = InStr(mudtProps.LookupSeq, "WH") > 0
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.UsesWhitaker"
End Property
Public Property Get LookupSequence(i As Integer) As String
    On Error GoTo ErrHandler
    LookupSequence = Mid(LookupSeq, (i * 2) - 1, 2)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.LookupSequence(i)", i
End Property
Public Property Get VATRate() As Double
    On Error GoTo ErrHandler
    VATRate = mudtProps.VATRate
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo ErrHandler
    VATRateF = PBKSPercentF(mudtProps.VATRate)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.VATRateF"
End Property
Public Property Get EnforceSections() As Boolean
    EnforceSections = False
End Property
Public Property Get LastStockTakeDate() As Date
    LastStockTakeDate = mudtProps.LastStockTakeDate
End Property
Public Property Get DefaultCurrencyDivisor() As Long
    DefaultCurrencyDivisor = 100
End Property
Public Property Get DefaultCurrencyFormatString() As String
    DefaultCurrencyFormatString = "R#,###.00"
End Property
Public Property Get LoyaltyClubTypeID() As Long
    On Error GoTo ErrHandler
    LoyaltyClubTypeID = mudtProps.LoyaltyClubTypeID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LoyaltyClubTypeID"
End Property
Public Property Get CustomerTypeDictID() As Long
    On Error GoTo ErrHandler
    CustomerTypeDictID = mudtProps.CustomerTypeDictID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CustomerTypeDictID"
End Property
Public Property Get InterestGroupsDictID() As Long
    On Error GoTo ErrHandler
    InterestGroupsDictID = mudtProps.CustomerIGDictID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.InterestGroupsDictID"
End Property
Public Property Get DefaultStoreID() As Long
    On Error GoTo ErrHandler
    DefaultStoreID = mudtProps.DefaultStoreID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStoreID"
End Property
Public Property Let DefaultStoreID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultStoreID = val
    mudtProps.IsDirty = True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStoreID(val)", val
End Property
Public Property Get DefaultStore() As a_Store
    On Error GoTo ErrHandler
    Set DefaultStore = Stores(DefaultStoreID & "k")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStore"
End Property

Public Function SetVATRate(val As String)
    On Error GoTo ErrHandler
Dim dblvatPercent As Double
    If mcolStack.Count = 0 Then Err.Raise 383
    SetVATRate = True
    If Trim$(val) = "" Then
        SetVATRate = 0
    ElseIf Not ConvertToDBL(val, dblvatPercent) Then
        SetVATRate = False
        Exit Function
    End If
    If (dblvatPercent <= 100) Then
        mobjValid.RuleBroken "VATRATE", False
    Else
        mobjValid.RuleBroken "VATRATE", True
    End If
    mudtProps.VATRate = dblvatPercent
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetVATRate(val)", val
End Function
Public Property Let VATRate(val As Double)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.VATRate(val)", val
End Property

Public Property Get UnallocatedPT() As Long
    On Error GoTo ErrHandler
    UnallocatedPT = mudtProps.UnallocatedPT
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.UnallocatedPT"
End Property
Public Property Let UnallocatedPT(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.UnallocatedPT = val
    mudtProps.IsDirty = True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.UnallocatedPT(val)", val
End Property
Public Function Reload()
    On Error GoTo ErrHandler

    Dim objPersist As a_c_Configuration_P
    
    If mcolStack.Count > 0 Then Err.Raise 445
    If mudtProps.IsNew Then Err.Raise 445
    ' code to load the object goes here
    Set objPersist = New a_c_Configuration_P
    SetSuperState objPersist.Fetch()
    Set objPersist = Nothing
    Set tlSections = New z_TextList
    Set tlProductTypes = New z_TextList
    tlSections.Load ltDictionary, , dtCategory
    tlProductTypes.Load ltProductTypeActive
    LoadStores_tl ""
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration.ReLoad"
End Function

Public Property Get BookfindStatus() As Integer
    BookfindStatus = iBookfindStatus
End Property
Public Sub SetBookfindStatus(val As Integer)
    iBookfindStatus = val
End Sub
Public Function Currencies() As ch_Currencies
    On Error GoTo ErrHandler
    Set Currencies = mcolCurrencies
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.Currencies"
    Exit Function
ErrHandler:
    ErrorIn "a_c_Configuration.Currencies"
End Function

Public Property Get DefaultCurrency() As a_Currency
    On Error GoTo ErrHandler
    If DefaultCurrencyID > 0 Then
        Set DefaultCurrency = Currencies(DefaultCurrencyID & "k")
    Else
        Set DefaultCurrency = Nothing
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.DefaultCurrency"
    Exit Property
ErrHandler:
    ErrorIn "a_c_Configuration.DefaultCurrency"
End Property
Public Property Get LocalCurrency() As a_Currency
    On Error GoTo ErrHandler
    If LocalCurrencyID > 0 Then
        Set LocalCurrency = Currencies(LocalCurrencyID & "k")
    Else
        Set LocalCurrency = Nothing
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.LocalCurrency"
    Exit Property
ErrHandler:
    ErrorIn "a_c_Configuration.LocalCurrency"
End Property

Public Property Get DefaultCurrencyID() As Long
    On Error GoTo ErrHandler
    DefaultCurrencyID = mudtProps.DefaultCurrencyID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCurrencyID"
End Property
Public Property Let DefaultCurrencyID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultCurrencyID = val
    mudtProps.IsDirty = True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCurrencyID(val)", val
End Property

Public Property Get LocalCurrencyID() As Long
    On Error GoTo ErrHandler
    LocalCurrencyID = mudtProps.LocalCurrencyID
    Exit Property
ErrHandler:
    ErrorIn "a_c_Configuration.LocalCurrencyID"
End Property
Public Property Let LocalCurrencyID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LocalCurrencyID = val
    mudtProps.IsDirty = True
    Exit Property
ErrHandler:
    ErrorIn "a_c_Configuration.LocalCurrencyID(val)", val
End Property



