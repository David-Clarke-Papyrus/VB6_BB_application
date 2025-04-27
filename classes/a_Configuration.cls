VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mudtProps As ConfigProps
Private mcolStack As Collection
Private mcolCompanies As ch_Companies
Private mcolCurrencies As ch_Currencies
Private mcolStores As ch_Store
Private mcolStaff As ch_Staff
Private mcolDC As ch_DocumentControl
Private mcolRRs As ch_RoundingRule
Private mcolBICs As c_BICCodes
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private oDefaultCompany As a_Company
Private tlCatalogueHeadings As z_TextList
Private tlSections As z_TextList
Private tlReturnReasons As z_TextList
Private tlCOActions As z_TextList
Private tlSections_Short As z_TextList  'without '<ALL>'
Private tlMultibuys As z_TextList
Private tlProductTypes As z_TextList
Private tlProductTypes_Short As z_TextList  'without '<ALL>'
Private tlCustomerTypesActive As z_TextList
Private tlWorkstations As z_TextList
Private tlPrinters As z_TextList
Private tlDictSystem As z_TextList
Private tlProductStatus As z_TextList
Private mIGLitLunchID As Long
Private mIGLaunchID As Long
Private mIGPromotionID As Long
Private mIGBookClubID As Long
Private mIGSaleID As Long
Private mWebExportID As Long
Private iBookfindStatus As Integer
Event Valid(oRS As String, Status As Boolean)

Private colClassors As Collection
Private strors As String
Public Property Get BookfindStatus() As Integer
    BookfindStatus = iBookfindStatus
End Property
Public Sub SetBookfindStatus(val As Integer)
    iBookfindStatus = val
End Sub

Public Property Get Sections() As z_TextList
    Set Sections = tlSections
End Property
Public Property Get Sections_Short() As z_TextList
    Set Sections_Short = tlSections_Short
End Property
Public Property Get Multibuys() As z_TextList
    Set Multibuys = tlMultibuys
End Property
Public Property Get ProductStatus() As z_TextList
    Set ProductStatus = tlProductStatus
End Property
Public Property Get ReturnReasons() As z_TextList
    Set ReturnReasons = tlReturnReasons
End Property
Public Property Get COActions() As z_TextList
    Set COActions = tlCOActions
End Property

Public Property Get ProductTypes() As z_TextList
    Set ProductTypes = tlProductTypes
End Property
Public Property Get ProductTypes_Short() As z_TextList
    Set ProductTypes_Short = tlProductTypes_Short
End Property
Public Property Get CustomerTypesActive() As z_TextList
    Set CustomerTypesActive = tlCustomerTypesActive
End Property
Public Property Get DictSystem() As z_TextList
    Set DictSystem = tlDictSystem
End Property

Public Property Get Workstations() As z_TextList
    Set Workstations = tlWorkstations
End Property
Public Property Get Printers() As z_TextList
    Set Printers = tlPrinters
End Property
Public Property Get DefaultPrinterID() As Long
Dim var As Variant
    DefaultPrinterID = 0
    For Each var In tlPrinters
        If var(2) Then
            DefaultPrinterID = tlPrinters.Key(var(0))
            Exit For
        End If
    Next
End Property
Public Property Get CatalogueHeadings() As z_TextList
    Set CatalogueHeadings = tlCatalogueHeadings
End Property
Public Sub RefreshPrinters()
    On Error GoTo errHandler
    Set tlPrinters = Nothing
    Set tlPrinters = New z_TextList
    tlPrinters.Load ltPrinters
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RefreshPrinters"
End Sub

Public Sub RefreshProductTypes()
    On Error GoTo errHandler
    Set tlProductTypes = Nothing
    Set tlProductTypes_Short = Nothing
    Set tlProductTypes = New z_TextList
    Set tlProductTypes_Short = New z_TextList
    tlProductTypes.Load ltProductTypeActive, , "<ALL>"
    tlProductTypes_Short.Load ltProductTypeActive
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RefreshProductTypes"
End Sub
Public Sub RefreshDictSystemTypes()
    On Error GoTo errHandler
    Set tlDictSystem = Nothing
    Set tlDictSystem = New z_TextList
    tlDictSystem.Load ltDictSystem
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RefreshProductTypes"
End Sub

Public Sub RefreshWorkstations()
    On Error GoTo errHandler
    Set tlWorkstations = Nothing
    Set tlWorkstations = New z_TextList
    tlWorkstations.Load ltWorkstations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RefreshWorkstations"
End Sub

Public Sub RefreshCatalogueHeadings()
    On Error GoTo errHandler
    Set tlCatalogueHeadings = Nothing
    Set tlCatalogueHeadings = New z_TextList
    tlCatalogueHeadings.Load ltCatalogueHeadings
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RefreshCatalogueHeadings"
End Sub
Public Sub RefreshSections()
    On Error GoTo errHandler
    Set tlSections = Nothing
    Set tlMultibuys = Nothing
    Set tlSections_Short = Nothing
    Set tlSections = New z_TextList
    Set tlSections_Short = New z_TextList
    Set tlMultibuys = New z_TextList
    tlSections.Load ltSectionsAll, , "<ALL>"
    tlSections_Short.Load ltSectionsActive
    tlMultibuys.Load ltMultibuys, , "<N/A>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RefreshSections"
End Sub

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As ConfigData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As ConfigData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetState(buffer)", buffer
End Sub

Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "Companies", mcolCompanies.GetSuperState
    .WriteProperty "Currencies", mcolCurrencies.GetSuperState
    .WriteProperty "Stores", mcolStores.GetSuperState
    .WriteProperty "Staff", mcolStaff.GetSuperState
    .WriteProperty "DCs", mcolDC.GetSuperState
    .WriteProperty "RRs", mcolRRs.GetSuperState
    .WriteProperty "BICs", mcolRRs.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.GetSuperState"
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
    mcolCompanies.SetSuperState .ReadProperty("Companies")
    mcolCurrencies.SetSuperState .ReadProperty("Currencies")
    mcolStores.SetSuperState .ReadProperty("Stores")
    mcolDC.SetSuperState .ReadProperty("DCs")
    mcolStaff.SetSuperState .ReadProperty("Staff")
    mcolRRs.SetSuperState .ReadProperty("RRS")
    mcolBICs.SetSuperState .ReadProperty("BICs")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetSuperState(buffer)", buffer
End Sub
Public Function Reload()
    On Error GoTo errHandler

    Dim objPersist As a_Configuration_P
    
    If mcolStack.Count > 0 Then Err.Raise 445
    If mudtProps.IsNew Then Err.Raise 445
    ' code to load the object goes here
    Set objPersist = New a_Configuration_P
    SetSuperState objPersist.Fetch()
    Set objPersist = Nothing
    Set tlSections = Nothing
    Set tlSections_Short = Nothing
    Set tlProductStatus = Nothing
    Set tlProductTypes = Nothing
    Set tlProductTypes_Short = Nothing
    Set tlMultibuys = Nothing
    Set tlCustomerTypesActive = Nothing
    Set tlDictSystem = Nothing
    Set tlWorkstations = Nothing
    Set tlPrinters = Nothing
    Set tlCatalogueHeadings = Nothing
    
    Set tlSections = New z_TextList
    Set tlSections_Short = New z_TextList
    Set tlProductStatus = New z_TextList
    Set tlProductTypes = New z_TextList
    Set tlProductTypes_Short = New z_TextList
    Set tlMultibuys = New z_TextList
    Set tlCustomerTypesActive = New z_TextList
    Set tlDictSystem = New z_TextList
    Set tlWorkstations = New z_TextList
    Set tlPrinters = New z_TextList
    Set tlCatalogueHeadings = New z_TextList
    tlSections.Load ltSectionsAll, , "<ALL>"
    tlSections_Short.Load ltSectionsActive
    tlMultibuys.Load ltMultibuys, , "<N/A>"
    tlProductTypes.Load ltProductTypeActive, , "<ALL>"
    tlProductTypes.Load ltProductTypeActive
    tlProductStatus.Load ltProductStatus
    tlCustomerTypesActive.Load ltCustomerTypeActive
    tlDictSystem.Load ltDictSystem
    tlCatalogueHeadings.Load ltCatalogueHeadings
    tlWorkstations.Load ltWorkstations
    tlPrinters.Load ltPrinters
    SetDictSystemIDs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ReLoad"
End Function
Public Sub ReloadCatHeads()
    Set tlCatalogueHeadings = Nothing
    Set tlCatalogueHeadings = New z_TextList
    tlCatalogueHeadings.Load ltCatalogueHeadings
End Sub
Public Sub ReloadCategories()
    Set tlSections = Nothing
    Set tlSections_Short = Nothing
    Set tlSections = New z_TextList
    Set tlSections_Short = New z_TextList
    tlSections.Load ltSectionsAll, , "<ALL>"
    tlSections_Short.Load ltSectionsActive

End Sub
Public Function Load() As Long
    On Error GoTo errHandler
Dim OpenResult As Integer
    Dim objPersist As a_Configuration_P
    
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    ' code to load the object goes here
    Set objPersist = New a_Configuration_P
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    SetSuperState objPersist.Fetch()
    Set objPersist = Nothing
    Set tlSections = Nothing
    Set tlSections_Short = Nothing
    Set tlProductStatus = Nothing
    Set tlProductTypes = Nothing
    Set tlProductTypes_Short = Nothing
    Set tlMultibuys = Nothing
    Set tlCustomerTypesActive = Nothing
    Set tlDictSystem = Nothing
    Set tlWorkstations = Nothing
    Set tlPrinters = Nothing
    Set tlCatalogueHeadings = Nothing
    Set objPersist = Nothing
    Set tlSections = New z_TextList
    Set tlReturnReasons = New z_TextList
    Set tlCOActions = New z_TextList
    Set tlProductStatus = New z_TextList
    Set tlSections_Short = New z_TextList
    Set tlMultibuys = New z_TextList
    Set tlProductTypes = New z_TextList
    Set tlProductTypes_Short = New z_TextList
    Set tlDictSystem = New z_TextList
    Set tlWorkstations = New z_TextList
    Set tlPrinters = New z_TextList
    Set tlCustomerTypesActive = New z_TextList
    Set tlCatalogueHeadings = New z_TextList
    tlSections.Load ltSectionsAll, , "<ALL>"
    tlReturnReasons.Load ltSRAll
    tlCOActions.Load ltCOActionCode
    tlSections_Short.Load ltSectionsActive
    tlMultibuys.Load ltMultibuys, , "<N/A>"
    tlProductTypes.Load ltProductTypeActive, , "<ALL>"
    tlProductTypes_Short.Load ltProductTypeActive
    tlDictSystem.Load ltDictSystem
    tlCatalogueHeadings.Load ltCatalogueHeadings
    tlWorkstations.Load ltWorkstations
    tlCustomerTypesActive.Load ltCustomerTypeActive
    tlProductStatus.Load ltProductStatus
    
    tlPrinters.Load ltPrinters
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    SetDictSystemIDs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Load"
End Function
Private Sub SetDictSystemIDs()
Dim i As Integer
    For i = 1 To tlDictSystem.Count
        If UCase(tlDictSystem.ItemByOrdinalIndex(i)) = "LA" Then
            mIGLaunchID = tlDictSystem.Key("LA")
        End If
        If UCase(tlDictSystem.ItemByOrdinalIndex(i)) = "LL" Then
            mIGLitLunchID = tlDictSystem.Key("LL")
        End If
        If UCase(tlDictSystem.ItemByOrdinalIndex(i)) = "PR" Then
            mIGPromotionID = tlDictSystem.Key("PR")
        End If
        If UCase(tlDictSystem.ItemByOrdinalIndex(i)) = "SA" Then
            mIGSaleID = tlDictSystem.Key("SA")
        End If
        If UCase(tlDictSystem.ItemByOrdinalIndex(i)) = "BC" Then
            mIGBookClubID = tlDictSystem.Key("BC")
        End If
        If UCase(tlDictSystem.ItemByOrdinalIndex(i)) = "WEB" Then
            mWebExportID = tlDictSystem.Key("WEB")
        End If
    Next
End Sub
Public Function Currencies() As ch_Currencies
    On Error GoTo errHandler
    Set Currencies = mcolCurrencies
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Currencies"
End Function
Public Function Stores() As ch_Store
    On Error GoTo errHandler
    Set Stores = mcolStores
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Stores"
End Function
Public Function Staff() As ch_Staff
    On Error GoTo errHandler
    Set Staff = mcolStaff
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Staff"
End Function
Public Function DocumentControls() As ch_DocumentControl
    On Error GoTo errHandler
    Set DocumentControls = mcolDC
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DocumentControl"
End Function
Public Function RoundingRules() As ch_RoundingRule
    On Error GoTo errHandler
    Set RoundingRules = mcolRRs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.RoundingRules"
End Function
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mcolCompanies = New ch_Companies
    Set mcolCurrencies = New ch_Currencies
    Set mcolStores = New ch_Store
    Set mcolStaff = New ch_Staff
    Set mcolDC = New ch_DocumentControl
    Set mcolRRs = New ch_RoundingRule
    Set mcolBICs = New c_BICCodes
    Set mobjValid = New z_BrokenRules
    Set colClassors = New Collection
    LoadClassorsCollection
  
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    
    Set mcolStack = Nothing
    Set mcolCompanies = Nothing
    Set mcolCurrencies = Nothing
    Set mcolStores = Nothing
    Set mcolStaff = Nothing
    Set mcolDC = Nothing
    Set mcolRRs = Nothing
    Set mcolBICs = Nothing
    Set mobjValid = Nothing
    Set colClassors = Nothing
    Set tlWorkstations = Nothing
    Set tlProductTypes = Nothing
    Set tlProductTypes_Short = Nothing
    Set tlDictSystem = Nothing
    Set tlCustomerTypesActive = Nothing
    Set tlCatalogueHeadings = Nothing
    Set tlSections = Nothing
    Set tlSections_Short = Nothing
    Set tlPrinters = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Class_Terminate"
End Sub
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid VAT rate", "VATRate"
    colClassors.Add "Invalid customer order line allocation style", "COLALLOC"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LoadClassorsCollection"
End Sub
Public Function Companies() As ch_Companies
    On Error GoTo errHandler
    Set Companies = mcolCompanies
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Companies"
End Function
Public Function BICs() As c_BICCodes
    On Error GoTo errHandler
    Set BICs = mcolBICs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.BICs"
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
    ErrorIn "a_Configuration.TranslateErrors(pRawors)", pRawors
End Function
Public Property Get COTypesSupported() As Integer
    On Error GoTo errHandler
    COTypesSupported = mudtProps.COTypesSupported
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.COTypesSupported"
End Property
Public Property Get SupportsWants() As Boolean
    On Error GoTo errHandler
    SupportsWants = (COTypesSupported = 1 Or COTypesSupported = 3)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SupportsWants"
End Property
Public Property Get SupportsStandingOrders() As Boolean
    On Error GoTo errHandler
    SupportsStandingOrders = (COTypesSupported = 2 Or COTypesSupported = 3)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SupportsStandingOrders"
End Property
Public Property Let COTypesSupported(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.COTypesSupported Then
        mudtProps.COTypesSupported = val
        mudtProps.IsDirty = True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.COTypesSupported(val)", val
End Property

Private Sub mobjValid_BrokenRule(pors As String)
    On Error GoTo errHandler
    strors = TranslateErrors(pors)
    RaiseEvent Valid(TranslateErrors(pors), IsValid)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.mobjValid_BrokenRule(pors)", pors
End Sub
Private Sub mobjValid_RuleUnbroken(pors As String)
    On Error GoTo errHandler
    strors = TranslateErrors(pors)
    RaiseEvent Valid(TranslateErrors(pors), IsValid)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.mobjValid_RuleUnbroken(pors)", pors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    strors = ""
    RaiseEvent Valid("", True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.mobjValid_NoBrokenRules"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IsValid"
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mcolCompanies.IsDirty Or mcolCurrencies.IsDirty Or mcolStores.IsDirty Or mcolRRs.IsDirty Or mcolStaff.IsDirty Or mcolDC.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IsDirty"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IsEditing"
End Property

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolCompanies.BeginEdit
  mcolCurrencies.BeginEdit
  mcolStores.BeginEdit
  mcolStaff.BeginEdit
  mcolRRs.BeginEdit
  mcolDC.BeginEdit
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.BeginEdit"
End Sub
Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mcolCompanies.CancelEdit
  mcolCurrencies.CancelEdit
  mcolStores.CancelEdit
  mcolStaff.CancelEdit
  mcolRRs.CancelEdit
  mcolDC.CancelEdit
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CancelEdit"
End Sub
Public Sub ApplyEdit(pors As String)
    On Error GoTo errHandler
Dim objPersist As a_Configuration_P
Dim strStatus As String

    If mcolStack.Count = 0 Then Err.Raise 445
    mcolCompanies.ApplyEdit
    mcolCurrencies.ApplyEdit
    mcolStores.ApplyEdit
    mcolStaff.ApplyEdit
    mcolRRs.ApplyEdit
    mcolDC.ApplyEdit
    Set objPersist = New a_Configuration_P
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
    If strStatus > "" Then
        pors = Reportors
        BeginEdit
    Else
        pors = ""
        mudtProps.IsDirty = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ApplyEdit(pors)", pors
End Sub
Private Function Reportors() As String
    On Error GoTo errHandler
Dim oStaff As a_Staff
Dim strReportors As String
    
    strReportors = ""
    For Each oStaff In mcolStaff
        If oStaff.SQLSTatus <> "" Then
            strReportors = strReportors & vbCrLf & GetSQLMeaning(oStaff.SQLSTatus) & " in staff record: " & oStaff.StaffName & " (" & oStaff.StaffTel & ")"
        End If
    Next
    Reportors = Right(strReportors, Len(strReportors) - 2)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.Reportors"
End Function
Private Function GetSQLMeaning(pCode As String) As String
    On Error GoTo errHandler
    Select Case pCode
    Case "23000"
        GetSQLMeaning = "There is a disallowed duplicate value in a field"
    Case "22000"
        GetSQLMeaning = "A value is missing in a field"
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.GetSQLMeaning(pCODE)", pCode
End Function
Public Property Get DefaultCompany() As a_Company
    On Error GoTo errHandler
    Set DefaultCompany = Companies(CStr(DefaultCOMPID))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCompany"
End Property
Public Property Get DefaultCurrency() As a_Currency
    On Error GoTo errHandler
    If DefaultCurrencyID > 0 Then
        Set DefaultCurrency = Currencies.Item(DefaultCurrencyID & "k")
    Else
        Set DefaultCurrency = Nothing
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCurrency"
End Property
Public Property Get LocalCurrency() As a_Currency
    On Error GoTo errHandler
    If LocalCurrencyID > 0 Then
        Set LocalCurrency = Currencies.Item(LocalCurrencyID & "k")
    Else
        Set LocalCurrency = Nothing
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LocalCurrency"
End Property
'Public Property Get DaysInWeek() As Integer
'    On Error GoTo ErrHandler
'    DaysInWeek = mudtProps.DaysInWeek
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.DaysInWeek"
'End Property
'Public Property Let DaysInWeek(val As Integer)
'    On Error GoTo ErrHandler
'    If val > 0 And val < 8 Then
'        mudtProps.DaysInWeek = val
'        mudtProps.IsDirty = True
'    End If
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.DaysInWeek(val)", val
'End Property

Public Property Get MinMU() As Integer
    On Error GoTo errHandler
    MinMU = mudtProps.MinMU
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.MinMU"
End Property
Public Property Let MinMU(val As Integer)
    On Error GoTo errHandler
    If val > 0 Then
        mudtProps.MinMU = val
        mudtProps.IsDirty = True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.MinMU(val)", val
End Property

Public Property Get DiscountVATDefault() As Boolean
    On Error GoTo errHandler
    DiscountVATDefault = FNB(mudtProps.DiscountVAT)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DiscountVATDefault"
End Property
Public Property Let DiscountVATDefault(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DiscountVAT = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DiscountVATDefault(val)", val
End Property

''Public Property Get DaysTillCompact() As Integer
''    DaysTillCompact = mudtProps.DaysTillCompact
''End Property
''Public Property Let DaysTillCompact(val As Integer)
''    If mcolStack.Count = 0 Then Err.Raise 383
''    If val > 0 And val < 8 Then
''        mudtProps.DaysTillCompact = val
''        mudtProps.IsDirty = True
''    End If
''End Property
'Public Property Get NextNominalUpdate() As Date
'    On Error GoTo ErrHandler
'    NextNominalUpdate = mudtProps.NextNominalUpdate
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.NextNominalUpdate"
'End Property
'Public Property Let NextNominalUpdate(val As Date)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.NextNominalUpdate = val
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.NextNominalUpdate(val)", val
'End Property
'Public Property Get OpSetsAuto() As Long
'    On Error GoTo ErrHandler
'    OpSetsAuto = mudtProps.OpSetsAuto
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.OpSetsAuto"
'End Property
'Public Property Let OpSetsAuto(val As Long)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.OpSetsAuto = val
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.OpSetsAuto(val)", val
'End Property
Public Property Get TransactionPrefix() As String
    On Error GoTo errHandler
    TransactionPrefix = FNS(mudtProps.TransactionPrefix)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TransactionPrefix"
End Property
Public Property Let TransactionPrefix(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TransactionPrefix = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TransactionPrefix(val)", val
End Property
'Public Property Get WSStart() As Date
'    On Error GoTo ErrHandler
'    WSStart = mudtProps.WSStart
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.WSStart"
'End Property
'Public Property Get MSStart() As Date
'    On Error GoTo ErrHandler
'    MSStart = mudtProps.MSStart
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.MSStart"
'End Property
Public Property Get TFRDiscount() As Long
    On Error GoTo errHandler
    TFRDiscount = mudtProps.TFRDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TFRDiscount"
End Property
Public Property Get TFRDiscountF() As String
    On Error GoTo errHandler
    TFRDiscountF = CStr(mudtProps.TFRDiscount) & "%"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TFRDiscountF"
End Property
Public Property Let TFRDiscount(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TFRDiscount = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TFRDiscount(val)", val
End Property
Public Property Get TFRDiscountAdj() As Long
    On Error GoTo errHandler
    TFRDiscountAdj = mudtProps.TFRDiscountAdj
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TFRDiscountAdj"
End Property
Public Property Get TFRDiscountAdjF() As String
    On Error GoTo errHandler
    TFRDiscountAdjF = mudtProps.TFRDiscountAdj & "%"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TFRDiscountAdjF"
End Property
Public Property Let TFRDiscountAdj(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TFRDiscountAdj = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.TFRDiscountAdj(val)", val
End Property
'Public Property Get CurrentMonth()
''Dim iBaseYear
' '   iBaseYear = gBaseYear
'    CurrentMonth = ((DatePart("yyyy", Date) - CLng(mudtProps.BaseYear)) * 12) + DatePart("m", Date)
'End Property
'Public Property Get CurrentWeek()
''Dim iBaseYear
''    iBaseYear = gBaseYear
'    CurrentWeek = (DatePart("yyyy", Date) - CLng(mudtProps.BaseYear)) * 52 + DatePart("ww", Date, vbMonday)
'End Property
'Public Property Get CurrentWeek2(pDate As Date)
''Dim iBaseYear
''    iBaseYear = gBaseYear
'    CurrentWeek2 = (DatePart("yyyy", pDate) - CLng(mudtProps.BaseYear)) * 52 + DatePart("ww", pDate, vbMonday)
'End Property
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = FNDBL(mudtProps.VATRate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.VATRateF"
End Property
Public Function SetVATRate(val As String)
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetVATRate(val)", val
End Function
Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.VATRate(val)", val
End Property
'Public Property Get VatNumber() As String
'    VatNumber = FNS(mudtProps.VatNumber)
'End Property
Public Property Get ConfirmBeforePrint() As Boolean
    On Error GoTo errHandler
    ConfirmBeforePrint = mudtProps.ConfirmBeforePrint
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ConfirmBeforePrint"
End Property
Public Property Let ConfirmBeforePrint(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ConfirmBeforePrint = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ConfirmBeforePrint(val)", val
End Property
Public Property Get PrintPrices() As Boolean
    On Error GoTo errHandler
    PrintPrices = mudtProps.PrintPrices
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.PrintPrices"
End Property
Public Property Let PrintPrices(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PrintPrices = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.PrintPrices(val)", val
End Property
Public Property Get IsVATRegion() As Boolean
    On Error GoTo errHandler
    IsVATRegion = mudtProps.VATRate > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IsVATRegion"
End Property
Public Property Let IsVATRegion(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.IsVATRegion = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IsVATRegion(val)", val
End Property

Public Property Get SupportsWORD() As Boolean
    On Error GoTo errHandler
    SupportsWORD = mudtProps.SupportsWORD
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SupportsWORD"
End Property
Public Property Let SupportsWORD(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SupportsWORD = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SupportsWORD(val)", val
End Property


Public Property Get PrintMemo() As Boolean
    On Error GoTo errHandler
    PrintMemo = mudtProps.PrintMemo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.PrintMemo"
End Property
Public Property Let PrintMemo(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PrintMemo = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.PrintMemo(val)", val
End Property
Public Property Get LocalCountryID() As Long
    On Error GoTo errHandler
    LocalCountryID = mudtProps.LocalCountryID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LocalCountryID"
End Property
Public Property Let LocalCountryID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LocalCountryID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LocalCountryID(val)", val
End Property

'-------
Public Property Get IGLaunchID() As Long
    On Error GoTo errHandler
    IGLaunchID = mIGLaunchID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IGLaunchID"
End Property
Public Property Get IGBookClubID() As Long
    On Error GoTo errHandler
    IGBookClubID = mIGBookClubID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IGBookClubID"
End Property

Public Property Get WebExportID() As Long
    On Error GoTo errHandler
    WebExportID = mWebExportID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.WebExportID"
End Property
Public Property Get IGSaleID() As Long
    On Error GoTo errHandler
    IGSaleID = mIGSaleID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IGSaleID"
End Property
Public Property Get IGPromotionID() As Long
    On Error GoTo errHandler
    IGPromotionID = mIGPromotionID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IGPromotionID"
End Property
Public Property Get IGLitLunchID() As Long
    On Error GoTo errHandler
    IGLitLunchID = mIGLitLunchID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.IGLitLunchID"
End Property
Public Property Get DefaultCurrencyID() As Long
    On Error GoTo errHandler
    DefaultCurrencyID = mudtProps.DefaultCurrencyID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCurrencyID"
End Property
Public Property Let DefaultCurrencyID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultCurrencyID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCurrencyID(val)", val
End Property
Public Property Get LocalCurrencyID() As Long
    On Error GoTo errHandler
    LocalCurrencyID = mudtProps.LocalCurrencyID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LocalCurrencyID"
End Property
Public Property Let LocalCurrencyID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LocalCurrencyID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LocalCurrencyID(val)", val
End Property
Public Property Get DefaultCOMPID() As Long
    On Error GoTo errHandler
    DefaultCOMPID = mudtProps.DefaultCOMPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCOMPID"
End Property
Public Property Let DefaultCOMPID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultCOMPID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCOMPID(val)", val
End Property
Public Property Get DefaultStoreID() As Long
    On Error GoTo errHandler
    DefaultStoreID = mudtProps.DefaultStoreID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStoreID"
End Property
Public Property Let DefaultStoreID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultStoreID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStoreID(val)", val
End Property
Public Property Get DefaultStore() As a_Store
    On Error GoTo errHandler
    Set DefaultStore = Stores.FindStoreByID(mudtProps.DefaultStoreID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStore"
End Property

Public Property Get BillToStoreID() As Long
    On Error GoTo errHandler
    BillToStoreID = mudtProps.BillToStoreID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.BillToStoreID"
End Property
Public Property Let BillToStoreID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BillToStoreID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.BillToStoreID(val)", val
End Property
Public Property Get BillToStore() As a_Store
    On Error GoTo errHandler
    Set BillToStore = Stores.Item(mudtProps.BillToStoreID & "k")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultStore"
End Property




Public Property Get CSCustomerID() As Long
    On Error GoTo errHandler
    CSCustomerID = mudtProps.CSCustomerID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CSCustomerID"
End Property
Public Property Let CSCustomerID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CSCustomerID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CSCustomerID(val)", val
End Property
Public Property Get InterestGroupsDictID() As Long
    On Error GoTo errHandler
    InterestGroupsDictID = mudtProps.InterestGroupsDictID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CustomerTypeDictID"
End Property
Public Property Get SectionTypeDictID() As Long
    On Error GoTo errHandler
    SectionTypeDictID = mudtProps.SectionTypeDictID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CustomerTypeDictID"
End Property

Public Property Get CustomerTypeDictID() As Long
    On Error GoTo errHandler
    CustomerTypeDictID = mudtProps.CustomerTypeDictID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CustomerTypeDictID"
End Property
Public Property Get DefaultPT() As Long
    On Error GoTo errHandler
    DefaultPT = FNN(mudtProps.DefaultPT)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultPT"
End Property
Public Property Let DefaultPT(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultPT = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultPT(val)", val
End Property
Public Property Get DefaultSection() As Long
    On Error GoTo errHandler
    DefaultSection = mudtProps.DefaultSection
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultSection"
End Property
Public Property Let DefaultSection(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultSection = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultSection(val)", val
End Property

Public Property Get DefaultCT() As Long
    On Error GoTo errHandler
    DefaultCT = mudtProps.DefaultCT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCT"
End Property
Public Property Let DefaultCT(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DefaultCT = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.DefaultCT(val)", val
End Property
Public Property Get LoyaltyClubTypeID() As Long
    On Error GoTo errHandler
    LoyaltyClubTypeID = mudtProps.LoyaltyClubType
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LoyaltyClubTypeID"
End Property



Public Property Get BusinessCustomerTypeID() As Integer
    On Error GoTo errHandler
    BusinessCustomerTypeID = mudtProps.BusinessCustomerTypeID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.BusinessCustomerTypeID"
End Property
Public Property Get PrivateCustomerTypeID() As Integer
    On Error GoTo errHandler
    PrivateCustomerTypeID = mudtProps.PrivateCustomerTypeID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.PrivateCustomerTypeID"
End Property
Public Property Get BookClubCustomerTypeID() As Integer
    On Error GoTo errHandler
    BookClubCustomerTypeID = mudtProps.BookClubCustomerTypeID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.BookClubCustomerTypeID"
End Property

Public Property Get LastUpdateDate() As Date
    On Error GoTo errHandler
    LastUpdateDate = mudtProps.LastUpdateDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastUpdateDate"
End Property
Public Property Get LastUpdateDateF() As String
    On Error GoTo errHandler
    LastUpdateDateF = Format(mudtProps.LastUpdateDate, "dd/mm/yyyy HH:nn")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastUpdateDateF"
End Property

Public Property Let LastUpdateDate(pDate As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LastUpdateDate = pDate
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastUpdateDate(pDate)", pDate
End Property
Public Property Get LastWantsExportDate() As Date
    On Error GoTo errHandler
    LastWantsExportDate = mudtProps.LastWantsExportDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastWantsExportDate"
End Property
Public Property Let LastWantsExportDate(pDate As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LastWantsExportDate = pDate
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastWantsExportDate(pDate)", pDate
End Property
Public Property Get UsesBookfind() As Boolean
    On Error GoTo errHandler
    UsesBookfind = InStr(UCase(mudtProps.LookupSeq), "BF") > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.UsesBookfind"
End Property
Public Property Get EnforceSections() As Boolean
    On Error GoTo errHandler
    EnforceSections = FNB(mudtProps.EnforceSections)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EnforceSections"
End Property
Public Property Let EnforceSections(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.EnforceSections = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EnforceSections(val)", val
End Property

Public Property Get ShowProdsWithInstancesOnly() As Boolean
    On Error GoTo errHandler
    ShowProdsWithInstancesOnly = FNB(mudtProps.ShowProdsWithInstancesOnly)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ShowProdsWithInstancesOnly"
End Property
Public Property Let ShowProdsWithInstancesOnly(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowProdsWithInstancesOnly = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ShowProdsWithInstancesOnly(val)", val
End Property
Public Property Get LoyaltyNumberingStartAt() As Long
    On Error GoTo errHandler
    LoyaltyNumberingStartAt = FNN(mudtProps.LoyStartNumber)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LoyaltyNumberingStartAt"
End Property
Public Sub SetLoyaltyNumberingStartAt(val As String)
Dim Res As Long
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If ConvertToLng(val, Res) = True Then
        mudtProps.LoyStartNumber = Res
    Else
        mudtProps.LoyStartNumber = 0
    End If
    mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetLoyaltyNumberingStartAt(val)", val
End Sub
Public Property Get LoyaltyNumberingEndAt() As Long
    On Error GoTo errHandler
    LoyaltyNumberingEndAt = FNN(mudtProps.LoyEndNumber)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LoyaltyNumberingEndAt"
End Property
Public Sub SetLoyaltyNumberingEndAt(val As String)
Dim Res As Long
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If ConvertToLng(val, Res) = True Then
        mudtProps.LoyEndNumber = Res
    Else
        mudtProps.LoyEndNumber = 0
    End If
    mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetLoyaltyNumberingEndAt(val)", val
End Sub
Public Property Let StatementText(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StatementText = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.StatementText(val)", val
End Property
Public Property Get StatementAsAt() As Date
StatementAsAt = FND(mudtProps.StatementAsAt)
End Property

Public Property Get StatementText() As String
    On Error GoTo errHandler
    StatementText = FNS(mudtProps.StatementText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.StatementText"
End Property
Public Property Get OrderText() As String
    On Error GoTo errHandler
    OrderText = FNS(mudtProps.OrderText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.OrderText"
End Property
Public Property Let OrderText(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OrderText = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.OrderText(val)", val
End Property
'    EmailPOMsg As String * 500
'    EmailInvMsg As String * 500
'    EmailCNMsg As String * 500
'    EmailAPPMsg As String * 500
'    EmailQuoteMsg As String * 500
Public Property Get EmailAPPMsg() As String
    On Error GoTo errHandler
    EmailAPPMsg = FNS(mudtProps.EmailAPPMsg)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailAPPMsg"
End Property
Public Property Let EmailAPPMsg(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.EmailAPPMsg = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailAPPMsg(val)", val
End Property
Public Property Get EmailQuoteMsg() As String
    On Error GoTo errHandler
    EmailQuoteMsg = FNS(mudtProps.EmailQuoteMsg)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailQuoteMsg"
End Property
Public Property Let EmailQuoteMsg(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.EmailQuoteMsg = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailQuoteMsg(val)", val
End Property
Public Property Get EmailInvMsg() As String
    On Error GoTo errHandler
    EmailInvMsg = FNS(mudtProps.EmailInvMsg)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailInvMsg"
End Property
Public Property Let EmailInvMsg(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.EmailInvMsg = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailInvMsg(val)", val
End Property
Public Property Get EmailPOMsg() As String
    On Error GoTo errHandler
    EmailPOMsg = FNS(mudtProps.EmailPOMsg)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailPOMsg"
End Property
Public Property Let EmailPOMsg(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.EmailPOMsg = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EmailPOMsg(val)", val
End Property































'Public Property Get EmailPOMsg() As String
'    On Error GoTo errHandler
'    EmailPOMsg = FNS(mudtProps.EmailPOMsg)
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.EmailPOMsg"
'End Property
'Public Property Let EmailPOMsg(val As String)
'    On Error GoTo errHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.EmailPOMsg = val
'    mudtProps.IsDirty = True
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.EmailPOMsg(val)", val
'End Property
'


Public Property Get SalesOrderText() As String
    On Error GoTo errHandler
    SalesOrderText = FNS(mudtProps.SalesOrderText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SalesOrderText"
End Property
Public Property Let SalesOrderText(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SalesOrderText = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SalesOrderText(val)", val
End Property

Public Property Get InvoiceText() As String
    On Error GoTo errHandler
    InvoiceText = FNS(mudtProps.InvoiceText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.InvoiceText"
End Property
Public Property Let InvoiceText(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.InvoiceText = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.InvoiceText(val)", val
End Property
Public Property Get QuotationText() As String
    On Error GoTo errHandler
    QuotationText = FNS(mudtProps.QuotationText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.QuotationText"
End Property
Public Property Let QuotationText(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.QuotationText = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.QuotationText(val)", val
End Property





Public Property Get LastStockTakeDate() As Date
    On Error GoTo errHandler
    LastStockTakeDate = mudtProps.LastStockTakeDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastStockTakeDate"
End Property
Public Property Let LastStockTakeDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsDate(val) Then
        mudtProps.LastStockTakeDate = val
        mudtProps.IsDirty = True
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LastStockTakeDate(val)", val
End Property
Public Property Get AllowCopyInfo() As Boolean
    On Error GoTo errHandler
    AllowCopyInfo = mudtProps.AllowCopyInfo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.AllowCopyInfo"
End Property
Public Property Let AllowCopyInfo(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.AllowCopyInfo = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.AllowCopyInfo(val)", val
End Property
Public Property Get AntiquarianYN() As Boolean
    On Error GoTo errHandler
    AntiquarianYN = mudtProps.AntiquarianYN
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.AntiquarianYN"
End Property
Public Property Let AntiquarianYN(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.AntiquarianYN = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.AntiquarianYN(val)", val
End Property

Public Property Get ReorderPerCOL() As Boolean
    On Error GoTo errHandler
    ReorderPerCOL = mudtProps.ReorderPerCOL
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ReorderPerCOL"
End Property
Public Property Let ReorderPerCOL(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ReorderPerCOL = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ReorderPerCOL(val)", val
End Property

Public Property Get EnforceCOLRef() As Boolean
    On Error GoTo errHandler
    EnforceCOLRef = mudtProps.EnforceCOLRef
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EnforceCOLRef"
End Property
Public Property Let EnforceCOLRef(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.EnforceCOLRef = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.EnforceCOLRef(val)", val
End Property

Public Property Get AggregatePOs() As Boolean
    On Error GoTo errHandler
    AggregatePOs = mudtProps.AggregatePOs
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.AggregatePOs"
End Property
Public Property Let AggregatePOs(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.AggregatePOs = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.AggregatePOs(val)", val
End Property




Public Property Get CasualCustomersYN() As Boolean
    On Error GoTo errHandler
    CasualCustomersYN = mudtProps.CasualCustomersYN
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CasualCustomersYN"
End Property
Public Property Let CasualCustomersYN(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CasualCustomersYN = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CasualCustomersYN(val)", val
End Property
'Public Property Get COLAllocationStyle() As String
'    On Error GoTo errHandler
'    COLAllocationStyle = FNS(mudtProps.COLAllocationStyle)
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.COLAllocationStyle"
'End Property
'Public Property Let COLAllocationStyle(val As String)
'    On Error GoTo errHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.COLAllocationStyle = val
'    mudtProps.IsDirty = True
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.COLAllocationStyle(val)", val
'End Property
'Public Function SetCOLAllocationStyle(val As String) As Boolean
'    On Error GoTo errHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    If (val = "R" Or val = "S") Then
'        mobjValid.RuleBroken "COLALLOC", False
'    Else
'        mobjValid.RuleBroken "COLALLOC", True
'    End If
'
'    mudtProps.COLAllocationStyle = val
'    mudtProps.IsDirty = True
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.SetCOLAllocationStyle(val)", val
'End Function

Public Property Get OfferSignature() As String
    On Error GoTo errHandler
    OfferSignature = FNS(mudtProps.OfferSignature)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.OfferSignature"
End Property
Public Property Let OfferSignature(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OfferSignature = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.OfferSignature(val)", val
End Property

Public Property Get CreditorsContraAccount() As String
    CreditorsContraAccount = FNS(mudtProps.IE_CreditorsContraAccount)
End Property
Public Property Let CreditorsContraAccount(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.IE_CreditorsContraAccount = val
    mudtProps.IsDirty = True
End Property
Public Property Get DebtorsContraAccount() As String
    DebtorsContraAccount = FNS(mudtProps.IE_DebtorsContraAccount)
End Property
Public Property Let DebtorsContraAccount(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.IE_DebtorsContraAccount = val
    mudtProps.IsDirty = True
End Property
Public Property Get GLReference() As String
    GLReference = FNS(mudtProps.IE_GLReference)
End Property
Public Property Let GLReference(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.IE_GLReference = val
    mudtProps.IsDirty = True
End Property
Public Property Get AccountingApplicationName() As String
    AccountingApplicationName = FNS(mudtProps.IE_AccountingApplicationName)
End Property
Public Property Let AccountingApplicationName(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.IE_AccountingApplicationName = val
    mudtProps.IsDirty = True
End Property


Public Property Get LookupSeq() As String
    On Error GoTo errHandler
    LookupSeq = FNS(mudtProps.LookupSeq)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LookupSeq"
End Property
Public Function SetLookupSequence(val As String) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LookupSeq = UCase(val)
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetLookupSequence(val)", val
End Function
Public Property Get PrintingSettings() As String
    On Error GoTo errHandler
    PrintingSettings = FNS(mudtProps.PrintingSettings)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.PrintingSettings"
End Property
Public Function SetPrintingSettings(val As String) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PrintingSettings = val
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetPrintingSettings(val)", val
End Function
Public Property Get GFXNumber() As String
    On Error GoTo errHandler
    GFXNumber = FNS(mudtProps.GFXNumber) & String$(13 - Len(FNS(mudtProps.GFXNumber)), 48)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.GFXNumber"
End Property
Public Function SetGFXNumber(val As String) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.GFXNumber = val
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SetGFXNumber(val)", val
End Function

Public Property Get UsesWhitaker() As Boolean
    On Error GoTo errHandler
    UsesWhitaker = InStr(mudtProps.LookupSeq, "WH") > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.UsesWhitaker"
End Property
Public Property Get LookupSequence(i As Integer) As String
    On Error GoTo errHandler
    LookupSequence = Mid(LookupSeq, (i * 3) - 2, 2)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.LookupSequence(i)", i
End Property
Public Property Get ShowDeposit() As Boolean
    On Error GoTo errHandler
    ShowDeposit = mudtProps.ShowDeposit
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ShowDeposit"
End Property
Public Property Let ShowDeposit(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowDeposit = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.ShowDeposit(val)", val
End Property
Public Property Get NonBookYN() As Boolean
    On Error GoTo errHandler
    NonBookYN = mudtProps.NonBookYN
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.NonBookYN"
End Property
Public Property Let NonBookYN(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.NonBookYN = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.NonBookYN(val)", val
End Property


Public Property Get SupportsLoyaltyClub() As Boolean
    On Error GoTo errHandler
    SupportsLoyaltyClub = mudtProps.SupportsLoyaltyClub
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SupportsLoyaltyClub"
End Property
Public Property Let SupportsLoyaltyClub(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SupportsLoyaltyClub = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SupportsLoyaltyClub(val)", val
End Property





Public Property Get SignTransactions() As Boolean
    On Error GoTo errHandler
    SignTransactions = mudtProps.SignTransactions
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SignTransactions"
End Property
Public Property Let SignTransactions(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SignTransactions = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.SignTransactions(val)", val
End Property
Public Property Get CaptureDecimal() As Boolean
    On Error GoTo errHandler
    CaptureDecimal = mudtProps.CaptureDecimal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CaptureDecimal"
End Property
Public Property Let CaptureDecimal(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDecimal = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration.CaptureDecimal(val)", val
End Property
'Public Property Get UpdateWindowStart() As Date
'    On Error GoTo ErrHandler
'    UpdateWindowStart = mudtProps.UpdateWindowStart
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.UpdateWindowStart"
'End Property
'Public Property Let UpdateWindowStart(pDate As Date)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.UpdateWindowStart = pDate
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.UpdateWindowStart(pDate)", pDate
'End Property
'Public Function SetUpdateWindowStart(val As String) As Boolean
'    On Error GoTo ErrHandler
'Dim dteDate As Date
'    SetUpdateWindowStart = True
'    If Trim$(val) = "" Then
'        val = CDate(0)
'    ElseIf Not ConvertToDate(val, dteDate) Then
'        SetUpdateWindowStart = False
'        Exit Function
'    End If
'    mudtProps.UpdateWindowStart = dteDate
'    mudtProps.IsDirty = True
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.SetUpdateWindowStart(val)", val
'End Function
'Public Property Get UpdateWindowStartFormatted() As String
'    On Error GoTo ErrHandler
'    UpdateWindowStartFormatted = Format(mudtProps.UpdateWindowStart, "h:mm AM/PM")
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.UpdateWindowStartFormatted"
'End Property
'Public Property Get UpdateWindowEnd() As Date
'    On Error GoTo ErrHandler
'    UpdateWindowEnd = mudtProps.UpdateWindowEnd
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.UpdateWindowEnd"
'End Property
'Public Property Let UpdateWindowEnd(pDate As Date)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.UpdateWindowEnd = pDate
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.UpdateWindowEnd(pDate)", pDate
'End Property
'Public Function SetUpdateWindowEnd(val As String) As Boolean
'    On Error GoTo ErrHandler
'Dim dteDate As Date
'    SetUpdateWindowEnd = True
'    If Trim$(val) = "" Then
'        val = CDate(0)
'    ElseIf Not ConvertToDate(val, dteDate) Then
'        SetUpdateWindowEnd = False
'        Exit Function
'    End If
'    mudtProps.UpdateWindowEnd = dteDate
'    mudtProps.IsDirty = True
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.SetUpdateWindowEnd(val)", val
'End Function
'Public Property Get UpdateWindowEndFormatted() As String
'    On Error GoTo ErrHandler
'    UpdateWindowEndFormatted = Format(mudtProps.UpdateWindowEnd, "h:mm AM/PM")
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Configuration.UpdateWindowEndFormatted"
'End Property

Public Function OptionLoopStores(pCurrent As Long, pExcludeLocal As Boolean) As Long
Dim iNext As Integer

    iNext = pCurrent + 1

    If iNext > Stores.Count Then iNext = 1
    If pExcludeLocal Then
        If oPC.Configuration.Stores(iNext) Is oPC.Configuration.DefaultStore Then
            If iNext = 1 And Stores.Count = 1 Then
                iNext = 0
            Else
                iNext = iNext + 1
            End If
        End If
    End If
    OptionLoopStores = iNext
End Function


Public Function ChangePassword(pShortname As String, pOldPassword As String, pNewPassword As String, pFullName As String) As Boolean
Dim strfullname As String
Dim lngStaffID As Long
Dim str As String


    If Staff.GetLevel(Trim(pShortname) & Trim(pOldPassword), strfullname, lngStaffID) > 0 Then
        oPC.Configuration.BeginEdit
        If Staff.FindStaffByID(lngStaffID).SetPassword(pNewPassword) Then
            oPC.Configuration.ApplyEdit str
          '  oPC.Configuration.IsEditing
            pFullName = strfullname
            ChangePassword = True
        Else
            oPC.Configuration.CancelEdit
            ChangePassword = False
        End If
    Else
        ChangePassword = False
    End If
End Function
Public Sub DeletePrinter(pPrinterId As Long)
Dim oSQL As New z_SQL
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    oSQL.DeletePrinter pPrinterId
    RefreshPrinters
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
End Sub

