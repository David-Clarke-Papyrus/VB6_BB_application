VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ro_Product"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Copies"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Copy"
  Option Explicit

Private mudtProps As ProductProps
Private mudtSave As ProductProps
Private mcolCopies As ch_Copy
Private mcolWants As c_COLSPerPID
Private mcolStores As ch_StoreP
Private mOSSO As c_OSSOrder
Private mOSCO As c_OSCOrder
Private mOSAPP As c_OSAPPRO
Private mcolStack As Collection
Private tlCategories As z_TextList
Private tlCatalogueHeadings As z_TextList
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private WithEvents oProdCode As z_ProdCode
Attribute oProdCode.VB_VarHelpID = -1
Private mcolClassors As Collection
Event Valid(IsValid As Boolean, msg As String)
Event CodeToBeGenerated()
Event RedisplayCodes()
Private oDefaultCopy As a_Copy
''Dim oBF As a_BookFind

Private flgEditing As Boolean
Private Function GetState() As String
  Dim udtData As ProductData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
End Function

Private Sub SetState(buffer As String)
  Dim udtData As ProductData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub

Public Function GetSuperState() As String
Dim objPB As PropertyBag
  
    Set objPB = New PropertyBag
    objPB.WriteProperty "State", GetState
    objPB.WriteProperty "Copies", mcolCopies.GetSuperState
  '  objPB.WriteProperty "Wants", mcolWants.GetSuperState
    objPB.WriteProperty "Stores", mcolStores.GetSuperState
    GetSuperState = objPB.Contents
    Set objPB = Nothing
End Function

Public Sub SetSuperState(buffer As String)
Dim objPB As PropertyBag
Dim arBuffer() As Byte
  
    Set objPB = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    SetState objPB.ReadProperty("State")
 '   If mudtProps.KeepsCopies Then
        mcolCopies.SetSuperState objPB.ReadProperty("Copies")
        mcolWants.SetSuperState objPB.ReadProperty("Wants")
        mcolStores.SetSuperState objPB.ReadProperty("Stores")
 '   End If
    Set objPB = Nothing
End Sub


'Public Function Exists(pCODE As String) As Boolean
'Dim objPersist As a_Product_p
'    Set objPersist = New a_Product_p
'    Exists = objPersist.Exists(pCODE)
'    Set objPersist = Nothing
'End Function
Private Function PrepareTitle(pIn) As Variant
Dim strTmp
    If IsNull(pIn) Then
        PrepareTitle = "<Unknown>"
        GoTo EXIT_PrepareTItle
    End If
    If oPC.Configuration.SignTransactions Then
        If Left$(pIn, 4) = "The " Then
            PrepareTitle = Right$(pIn, Len(pIn) - 4)
        ElseIf Left$(pIn, 2) = "A " Then
            PrepareTitle = Right$(pIn, Len(pIn) - 2)
        ElseIf Left$(pIn, 3) = "An " Then
            PrepareTitle = Right$(pIn, Len(pIn) - 2)
        ElseIf Left$(pIn, 3) = "'n " Then
            PrepareTitle = Right$(pIn, Len(pIn) - 3)
        ElseIf Left$(pIn, 2) = "n " Then
            PrepareTitle = Right$(pIn, Len(pIn) - 2)
        Else
            PrepareTitle = pIn
        End If
    Else
        PrepareTitle = pIn
    End If
EXIT_PrepareTItle:
    Exit Function

End Function

Public Property Get Copies() As ch_Copy
    Set Copies = mcolCopies
End Property
Public Property Get Wants() As c_COLSPerPID
    Set Wants = mcolWants
End Property
Public Property Get Stores() As ch_StoreP
    Set Stores = mcolStores
End Property

Public Property Get IsEditing() As Boolean
    IsEditing = mcolStack.Count > 0
End Property
Public Sub BeginEdit()
    mcolCopies.BeginEdit
  '  mcolWants.BeginEdit
    mcolStores.BeginEdit
    mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mcolCopies.CancelEdit
'  mcolWants.CancelEdit
  mcolStores.CancelEdit
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub
'Private Sub Checkcodes()
'Dim oPC As z_ProdCode
'    Set oPC = New z_ProdCode
'    oPC.Load Me.Code, True
'
'End Sub
Public Sub ApplyEdit(Optional plResult As Long)
On Error GoTo ERR_Handler
Dim objPersist As a_Product_p
Dim oProdCode As z_ProdCode
Dim strTemp As String
    plResult = 0
    If mcolStack.Count = 0 Then Err.Raise 445
    
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        mudtProps.SpecialVat = True
    End If

    mcolCopies.ApplyEdit
'    mcolWants.ApplyEdit
    mcolStores.ApplyEdit
    Set objPersist = New a_Product_p
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject Trim$(mudtProps.ID)
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    strTemp = objPersist.Save(GetSuperState)
    If strTemp > "" Then
        mcolStack.Remove mcolStack.Count
        SetSuperState strTemp
        mudtProps.IsNew = False
    Else 'save failed
        mcolCopies.BeginEdit
    '    mcolWants.BeginEdit
        mcolStores.BeginEdit
        plResult = 99
        Set objPersist = Nothing
        GoTo EXIT_Handler
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
    plResult = 101
    MsgBox "Unknown or" & Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mcolCopies = New ch_Copy
    Set mcolWants = New c_COLSPerPID
    Set mcolStores = New ch_StoreP
    Set mOSSO = New c_OSSOrder
    Set mOSCO = New c_OSCOrder
    Set mOSAPP = New c_OSAPPRO

    Set mobjValid = New z_BrokenRules
    Set mcolClassors = New Collection
    Set oProdCode = New z_ProdCode
    Set tlCategories = New z_TextList
    Set tlCatalogueHeadings = New z_TextList
    tlCategories.Load ltProductType
    tlCatalogueHeadings.Load ltCatalogueHeadings
    LoadClassorsCollection
    mudtProps.IsNew = True
End Sub
Public Sub Initialize(pbKeepsCopies As Boolean)
    mudtProps.KeepsCopies = pbKeepsCopies
    mudtProps.SpecialVat = False
    mudtProps.VATRate = oPC.Configuration.VATRate
End Sub
Private Sub Class_Terminate()
    If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    
    Set mcolStack = Nothing
    Set mcolCopies = Nothing
    Set mcolWants = Nothing
    Set mcolStores = Nothing
    Set mOSSO = Nothing
    Set mOSCO = Nothing
    Set mOSAPP = Nothing

    Set mobjValid = Nothing
    Set mcolClassors = Nothing
    Set oProdCode = Nothing
End Sub

Public Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property
Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty Or mcolCopies.IsDirty
End Property
Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
  RaiseEvent Valid(False, oRS)
End Sub
Private Sub mobjValid_NoBrokenRules()
  RaiseEvent Valid(True, "")
End Sub
Private Sub mobjValid_Status(sRulesBroken As String)
    RaiseEvent Valid((sRulesBroken = ""), TranslateErrors(sRulesBroken))
End Sub
Public Property Get DefaultCopy() As a_Copy
    Set DefaultCopy = oDefaultCopy
End Property
Public Sub PrepareForReload()
    mudtProps.IsNew = True
End Sub
'Public Function Load(PID As String, pCatID As Long, Optional pCODE As String) As Long   ', Optional pbReload As Boolean
''Returns 0 if found on db or loaded and saved from Bookfind
''Returns 99 if an or occurs on saving from Bookfind
''Returns 98 if not found on db or Bookfind
'
'On Error GoTo ERR_Handler
'Dim strCode As String
'Dim lngResult As Long
'Dim lngFound As Long
'Dim lngPK As Long
'Dim iProdID As Long
'Dim iPos As Integer
'Dim tmpBuffer As String
'Dim tmp As String
'Dim objPersist As a_Product_p
'Dim i As Integer
'
'    strCode = FNS(pCODE)
'    If mcolStack.Count > 0 Then Err.Raise 445
'    If Not mudtProps.IsNew Then Err.Raise 445
'
'    Set objPersist = New a_Product_p
'START:
''option 1
'    lngFound = 99
'    If Len(PID) > 1 Then
'        tmp = objPersist.Fetch(lngResult, PID) 'returns 0 if found; 99 if not found
'        If tmp > "" Then
'            lngFound = 0
'            SetSuperState tmp
'            GoTo EXIT_Handler
'        Else
'            lngFound = 0
'            GoTo EXIT_Handler
'        End If
'    End If
'
''option 2
'    If IsISBN13(strCode) Then
'        tmpBuffer = objPersist.Fetch(lngResult, , , strCode)
'        If tmpBuffer > "" Then
'            SetSuperState tmpBuffer
'            lngFound = 0
'            GoTo EXIT_Handler
'        Else
'            lngFound = 99
'        End If
'        'Convert ISBN13 to ISBN10 if possible so that next search may work
'        strCode = ISBN13to10(strCode)
'    End If
'
''option 3
'    If IsISBN10(strCode) Or IsHashCode(strCode) Or IsPrivateCode(strCode) Then   'pCode s not an ISBN-13 or an EAN
'        tmpBuffer = objPersist.Fetch(lngResult, , strCode)
'        If tmpBuffer > "" Then
'            SetSuperState tmpBuffer
'            iPos = InStr(pCODE, "/")
'            If iPos > 0 Then 'There is a serial number
'                Set oDefaultCopy = Me.Copies.FindBySerial(Right(pCODE, Len(pCODE) - iPos))
'            End If
'            If oDefaultCopy Is Nothing Then
'                lngFound = 99
'                GoTo EXIT_Handler
'            Else
'                lngFound = 0
'            End If
'        Else
'            lngFound = 99
'        End If
'    End If
'
'
''we only get here if we have not found the record
'    If IsISBN13(FNS(pCODE)) Or IsISBN10(FNS(pCODE)) Then
'        i = 1
'            Do While oPC.Configuration.LookupSequence(i) > "" And lngResult <> 0
'             Select Case oPC.Configuration.LookupSequence(i)
'             Case "BF"
'                    Set oBF = New a_BookFind
'                    lngResult = oBF.FetchFromBF(oProdCode.code)
'                    If lngResult = 0 Then  'Found a record
'                        lngFound = 0
'                        LoadProductFromBF
'                        SetSuperState objPersist.Save(GetSuperState)
'                        PID = mudtProps.ID
'                        Set oBF = Nothing
'                        GoTo START
'                        mudtProps.IsDirty = True
'                        mudtProps.IsNew = False
'                    Else
'                        PID = 0 'product not on db and not on Bookfind
'                        lngFound = 99
'                    End If
'                    Set oBF = Nothing
'             Case "WH"
'             End Select
'            i = i + 1
'        Loop
'    End If
'    mOSSO.Load Me.PID
'    mOSCO.Load Me.PID
'    mOSAPP.Load Me.PID
'
'    Load = lngFound
'EXIT_Handler:
'    Set oProdCode = Nothing
'    Exit Function
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'End Function


Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub

'Public Function LookOnBookfind(pCODE As String) As Integer
'Dim iresult As Long
'    Set oBF = New a_BookFind
'    iresult = oBF.DestroyResultsList
'    iresult = oBF.FetchFromBF(pCODE)
'    If iresult > 0 Then
'        LoadProductFromBF
'    End If
'    LookOnBookfind = iresult
'    Set oBF = Nothing
'End Function
'Private Sub LoadProductFromBF()
'    mudtProps.code = oBF.code
'    mudtProps.Author = oBF.MainAuthor
'    mudtProps.Title = oBF.Title
'    mudtProps.SubTitle = oBF.SubTitle
'    mudtProps.Edition = oBF.Edition
'    mudtProps.Publisher = oBF.PublisherName
'    mudtProps.Note = oBF.Note
'    mudtProps.SeriesTitle = oBF.SeriesTitle
'    mudtProps.Availability = oBF.Availability
'    mudtProps.UKPrice = oBF.UKPrice
'    mudtProps.USPrice = oBF.USPrice
'    mudtProps.RRP = oBF.LocalPrice
'    mudtProps.PublicationDate = oBF.PublicationDate
'    mudtProps.Description = oBF.Description
'    mudtProps.MainSupplierName = oBF.MainSupplierName
'    mudtProps.BindingCode = oBF.BindingCode
'    mudtProps.BFClassification = oBF.BFClassification
'    MsgBox "Possibly problem assigning properties directly -specially prices"
'End Sub
Public Property Get PublicationDate() As String
    PublicationDate = Trim(mudtProps.PublicationDate)
End Property
Public Function SetPublicationDate(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPublicationDate = True
    If Len(val) > Len(mudtProps.PublicationDate) Then
        Err.Raise 384
    End If
    mudtProps.PublicationDate = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
End Function
Public Property Get PublicationPlace() As String
    PublicationPlace = Trim(mudtProps.PublicationPlace)
End Property
Public Function SetPublicationPlace(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPublicationPlace = True
    If Len(val) > Len(mudtProps.PublicationPlace) Then
        Err.Raise 384
    End If
    mudtProps.PublicationPlace = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
End Function
Public Property Get Edition() As String
    Edition = Trim(mudtProps.Edition)
End Property
Public Function SetEdition(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetEdition = True
    If Len(val) > Len(mudtProps.Edition) Then
        Err.Raise 384
    End If
    mudtProps.Edition = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
End Function

Public Property Get PID() As String
    PID = FNS(mudtProps.ID)
End Property

Public Property Get Description() As String
    Description = Trim$(mudtProps.Description)
End Property
Public Function SetDescription(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetDescription = True
    If Len(val) > Len(mudtProps.Description) Then
        Err.Raise 384
    End If
    mudtProps.Description = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
End Function

Public Property Get Note() As String
    Note = Trim$(mudtProps.Note)
End Property
Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetNote = True
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    mudtProps.Note = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
End Function
Public Property Get code() As String
    code = FNS(mudtProps.code)
End Property
Public Property Get CodeF() As String
    CodeF = FNS(mudtProps.CodeF)
End Property
Public Property Get BottomOfDocument() As Boolean
    BottomOfDocument = (Right(mudtProps.code, 2) = "_B")
End Property

'Public Function GenerateCode()
'Dim oCodeGenerator As z_Code
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetCode = True
'    If Len(val) > Len(mudtProps.Code) Then
'        Err.Raise 384
'    End If
'    mudtProps.Code = val
'    oProdcode.Load val, True
'    SetCode = ValidateObject("CODE")
'    If EAN = "" Then
'        SetEAN oProdcode.EAN
'    End If
'    mudtProps.IsDirty = True
'    If val = "#" Then RaiseEvent CodeToBeGenerated
'    RaiseEvent RedisplayCodes
'
'End Function
Public Function SetCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetCode = True
    If Len(val) > Len(mudtProps.code) Then
        Err.Raise 384
    End If
    mudtProps.code = val
    SetCode = ValidateObject("CODE")
'    If EAN = "" Then
'        SetEAN oProdCode.EAN
'    End If
    mudtProps.IsDirty = True
    If val = "#" Then RaiseEvent CodeToBeGenerated
    RaiseEvent RedisplayCodes
End Function
Public Property Get EAN() As String
    EAN = FNS(mudtProps.EAN)
End Property
Public Function SetEAN(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetEAN = True
    If Len(val) > Len(mudtProps.EAN) Then
        Err.Raise 384
    End If
    mudtProps.EAN = val
    SetEAN = ValidateObject("CODE")
    mudtProps.IsDirty = True
    RaiseEvent RedisplayCodes
End Function
Public Property Get TitleAuthorPublisher() As String
Dim tmp As String

    tmp = Trim$(mudtProps.Title)
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Author)) > 0, "(" & Trim$(mudtProps.Author) & ")", "")
    End If
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Publisher)) > 0, "(" & Trim$(mudtProps.Publisher) & ")", "")
    End If
    TitleAuthorPublisher = tmp
End Property
Public Property Get BIC() As String
    BIC = Trim$(mudtProps.BIC)
End Property
Public Function SetBIC(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetBIC = True
    If Len(val) > Len(mudtProps.BIC) Then
        Err.Raise 384
    End If
    mudtProps.BIC = val
    SetBIC = ValidateObject("BIC")
    mudtProps.IsDirty = True
End Function

'Public Property Get ShelfPosition() As String
'    ShelfPosition = Trim$(mudtProps.ShelfPosition)
'End Property
'Public Function SetShelfPosition(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    val = Trim(val)
'    SetShelfPosition = True
'    If Len(val) > Len(mudtProps.ShelfPosition) Then
'        Err.Raise 384
'    End If
'    mudtProps.ShelfPosition = val
'    SetShelfPosition = ValidateObject("ShelfPosition")
'    mudtProps.IsDirty = True
'End Function

Public Property Get Comment() As String
    Comment = Trim$(mudtProps.Comment)
End Property
Public Function SetComment(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetComment = True
    If Len(val) > Len(mudtProps.Comment) Then
        Err.Raise 384
    End If
    mudtProps.Comment = val
    SetComment = ValidateObject("Comment")
    mudtProps.IsDirty = True
End Function

Public Property Get Summary() As String
    Summary = Trim$(mudtProps.Summary)
End Property
Public Function SetSummary(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetSummary = True
    If Len(val) > Len(mudtProps.Summary) Then
        Err.Raise 384
    End If
    mudtProps.Summary = val
    SetSummary = ValidateObject("Summary")
    mudtProps.IsDirty = True
End Function


Public Property Get BindingCode() As String
    BindingCode = Trim$(mudtProps.BindingCode)
End Property
Public Function SetBindingCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetBindingCode = True
    If Len(val) > Len(mudtProps.BindingCode) Then
        Err.Raise 384
    End If
    mudtProps.BindingCode = val
    SetBindingCode = ValidateObject("BindingCode")
    mudtProps.IsDirty = True
End Function


Public Property Get Title() As String
    Title = Trim$(mudtProps.Title)
End Property
Public Function SetTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetTitle = True
    If Len(val) > Len(mudtProps.Title) Then
        Err.Raise 384
    End If
    mudtProps.Title = val
    SetTitle = ValidateObject("TITLE")
    mudtProps.IsDirty = True
End Function
Public Property Get SubTitle() As String
    SubTitle = Trim(mudtProps.SubTitle)
End Property
Public Function SetSubTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetSubTitle = True
    If Len(val) > Len(mudtProps.SubTitle) Then
        Err.Raise 384
    End If
    mudtProps.SubTitle = val
    SetSubTitle = ValidateObject("SUBTITLE")
    mudtProps.IsDirty = True
End Function
Public Property Get FlagText() As String
    FlagText = Trim(mudtProps.FlagText)
End Property
Public Function SetFlagtext(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetFlagtext = True
    If Len(val) > Len(mudtProps.FlagText) Then
        Err.Raise 384
    End If
    mudtProps.FlagText = val
    SetFlagtext = ValidateObject("FLAGTEXT")
    mudtProps.IsDirty = True
End Function

Public Property Get CategoryID() As Long
    CategoryID = mudtProps.CategoryID
End Property
Public Property Get CategoryName() As String
    CategoryName = tlCategories.Item(mudtProps.CategoryID)
End Property
Public Function SetCategoryID(val As Long) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCategoryID = True
    mudtProps.CategoryID = val
    mudtProps.IsDirty = True
End Function
Public Property Get CatalogueheadingID() As Long
    CatalogueheadingID = mudtProps.CatalogueheadingID
End Property
Public Function SetCatalogueheadingID(val As Long) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCatalogueheadingID = True
    mudtProps.CatalogueheadingID = val
 '   SetCategoryID = ValidateObject(Val, "CATEGORYID")
    mudtProps.IsDirty = True
End Function
Public Property Get CatalogueHeadingName() As String
    CatalogueHeadingName = tlCatalogueHeadings.Item(mudtProps.CatalogueheadingID)
End Property

Public Property Get BFClassification() As String
    BFClassification = mudtProps.BFClassification
End Property

Public Property Get PriceLastDelivered() As Long
    PriceLastDelivered = mudtProps.PriceLastDelivered
End Property
Public Property Get PriceLastDeliveredF() As String
    If mudtProps.PriceLastDelivered = 0 Then
        PriceLastDeliveredF = ""
    Else
        PriceLastDeliveredF = Format(mudtProps.PriceLastDelivered / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Function SetPriceLastDelivered(val As String) As Boolean
Dim lngPriceLastDelivered As Integer
    SetPriceLastDelivered = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngPriceLastDelivered = 0
    ElseIf (Not ConvertToInt(val, lngPriceLastDelivered)) Then
        SetPriceLastDelivered = False
        Exit Function
    End If
    mudtProps.PriceLastDelivered = lngPriceLastDelivered
    SetPriceLastDelivered = ValidateObject("PriceLastDelivered")
    SetDirty True
End Function

Public Property Get PriceLastOrdered() As Long
    PriceLastOrdered = mudtProps.PriceLastOrdered
End Property
Public Property Get PriceLastOrderedF() As String
    If mudtProps.PriceLastOrdered = 0 Then
        PriceLastOrderedF = ""
    Else
        PriceLastOrderedF = Format(mudtProps.PriceLastOrdered / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If

End Property
Public Function SetPriceLastOrdered(val As String) As Boolean
Dim lngPriceLastOrdered As Integer
    SetPriceLastOrdered = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngPriceLastOrdered = 0
    ElseIf (Not ConvertToInt(val, lngPriceLastOrdered)) Then
        SetPriceLastOrdered = False
        Exit Function
    End If
    mudtProps.PriceLastOrdered = lngPriceLastOrdered
    SetPriceLastOrdered = ValidateObject("PriceLastOrdered")
    SetDirty True
End Function

Public Property Let DateLastDelivered(val As Date)
    mudtProps.DateLastDelivered = val
End Property
Public Property Get DateLastDelivered() As Date
    DateLastDelivered = mudtProps.DateLastDelivered
End Property
Public Property Get DateLastDeliveredF() As String
    If mudtProps.DateLastDelivered = CDate(0) Then
        DateLastDeliveredF = ""
    Else
        DateLastDeliveredF = Format(mudtProps.DateLastDelivered, "dd/mm/yyyy")
    End If
End Property

Public Function SetDateLastDelivered(val As String) As Boolean
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDateLastDelivered = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDateLastDelivered = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateLastDelivered", False
    Else
        mobjValid.RuleBroken "DateLastDelivered", True
    End If
    mudtProps.DateLastDelivered = dteDate
    mudtProps.IsDirty = True
End Function
Public Property Let DateLastOrdered(val As Date)
    mudtProps.DateLastOrdered = val
End Property
Public Property Get DateLastOrdered() As Date
    DateLastOrdered = mudtProps.DateLastOrdered
End Property
Public Property Get DateLastOrderedF() As String
    If mudtProps.DateLastOrdered = CDate(0) Then
        DateLastOrderedF = ""
    Else
        DateLastOrderedF = Format(mudtProps.DateLastOrdered, "dd/mm/yyyy")
    End If
End Property

Public Function SetDateLastOrdered(val As String) As Boolean
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDateLastOrdered = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDateLastOrdered = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateLastOrdered", False
    Else
        mobjValid.RuleBroken "DateLastOrdered", True
    End If
    mudtProps.DateLastOrdered = dteDate
    mudtProps.IsDirty = True
End Function

Public Property Get QtyLastDelivered() As Long
    QtyLastDelivered = mudtProps.QtyLastDelivered
End Property
Public Property Get QtyLastDeliveredF() As String
    If mudtProps.QtyLastDelivered = 0 Then
        QtyLastDeliveredF = ""
    Else
        QtyLastDeliveredF = mudtProps.QtyLastDelivered
    End If
End Property
Friend Property Let QtyLastDelivered(val As Long)
    mudtProps.QtyLastDelivered = val
End Property
Public Function SetQtyLastDelivered(val As String) As Boolean
Dim lngQty As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetQtyLastDelivered = True
    If Trim$(val) = "" Then
        lngQty = 0
    ElseIf Not ConvertToLng(val, lngQty) Then
        SetQtyLastDelivered = False
        Exit Function
    End If
    If (lngQty <= 1) Then
        mobjValid.RuleBroken "QtyLastDelivered", False
    Else
        mobjValid.RuleBroken "QtyLastDelivered", True
    End If
    mudtProps.QtyLastDelivered = lngQty
    mudtProps.IsDirty = True

End Function
Public Property Get QtyLastOrdered() As Long
    QtyLastOrdered = mudtProps.QtyLastOrdered
End Property
Public Property Get QtyLastOrderedF() As String
    If mudtProps.QtyLastDelivered = 0 Then
        QtyLastOrderedF = ""
    Else
        QtyLastOrderedF = mudtProps.QtyLastOrdered
    End If
End Property
Friend Property Let QtyLastOrdered(val As Long)
    mudtProps.QtyLastOrdered = val
End Property
Public Function SetQtyLastOrdered(val As String) As Boolean
Dim lngQty As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetQtyLastOrdered = True
    If Trim$(val) = "" Then
        lngQty = 0
    ElseIf Not ConvertToLng(val, lngQty) Then
        SetQtyLastOrdered = False
        Exit Function
    End If
    If (lngQty <= 1) Then
        mobjValid.RuleBroken "QtyLastOrdered", False
    Else
        mobjValid.RuleBroken "QtyLastOrdered", True
    End If
    mudtProps.QtyLastOrdered = lngQty
    mudtProps.IsDirty = True

End Function


Public Property Get RRP() As Long
    RRP = mudtProps.RRP
End Property
Public Property Get RRPF() As String
    RRPF = Format(mudtProps.RRP / 100, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Function SetRRP(val As String) As Boolean
Dim lngRRP As Integer
    SetRRP = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngRRP = 0
    ElseIf (Not ConvertToInt(val, lngRRP)) Then
        SetRRP = False
        Exit Function
    End If
    mudtProps.RRP = lngRRP
    SetRRP = ValidateObject("RRP")
    mudtProps.IsDirty = True
End Function
Public Property Get SP() As Currency
    SP = mudtProps.SP
End Property
Public Property Get SPF() As String
    SPF = Format(mudtProps.SP / 100, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Function SetSP(val As String) As Boolean
Dim lngSP As Integer
    SetSP = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngSP = 0
    ElseIf (Not ConvertToInt(val, lngSP)) Then
        SetSP = False
        Exit Function
    End If
    mudtProps.SP = lngSP
    SetSP = ValidateObject("SP")
    mudtProps.IsDirty = True
End Function
Public Property Get Cost() As Currency
    Cost = mudtProps.Cost
End Property
Public Property Get CostF() As String
    CostF = Format(mudtProps.Cost / 100, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Function SetCost(val As String) As Boolean
Dim lngCost As Integer
    SetCost = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngCost = 0
    ElseIf (Not ConvertToInt(val, lngCost)) Then
        SetCost = False
        Exit Function
    End If
    mudtProps.Cost = lngCost
    SetCost = ValidateObject("COST")
    mudtProps.IsDirty = True
End Function

'Public Property Get BindingCode() As String
'    BindingCode = mudtProps.BindingCode
'End Property
'Public Function SetBindingCode(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    val = Trim(val)
'    SetBindingCode = True
'    If Len(val) > Len(mudtProps.BindingCode) Then
'        Err.Raise 384
'    End If
'    mudtProps.BindingCode = val
'  '  SetSubTitle = ValidateObject(Val, "SUBTITLE")
'    mudtProps.IsDirty = True
'End Function


Public Property Get UKPrice() As Currency
    UKPrice = mudtProps.UKPrice
End Property
Public Function SetUKPrice(val As String) As Boolean
Dim lngUKPrice As Integer
    SetUKPrice = True
    If Trim$(val) = "" Then
        lngUKPrice = 0
    ElseIf (Not ConvertToInt(val, lngUKPrice)) Then
        SetUKPrice = False
        Exit Function
    End If
    mudtProps.UKPrice = lngUKPrice
    SetUKPrice = ValidateObject("UKPrice")
    mudtProps.IsDirty = True
End Function

Public Property Get USPrice() As Currency
    USPrice = mudtProps.USPrice
End Property
Public Function SetUSPrice(val As String) As Boolean
Dim lngUSPrice As Integer
    If mcolStack.Count = 0 Then Err.Raise 383
    SetUSPrice = True
    If Trim$(val) = "" Then
        lngUSPrice = 0
    ElseIf (Not ConvertToInt(val, lngUSPrice)) Then
        SetUSPrice = False
        Exit Function
    End If
    mudtProps.USPrice = lngUSPrice
    SetUSPrice = ValidateObject("USPrice")
    mudtProps.IsDirty = True
End Function

Public Property Get SeriesTitle() As String
    SeriesTitle = mudtProps.SeriesTitle
End Property
Public Function SetSeriesTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetSeriesTitle = True
    If Len(val) > Len(mudtProps.SeriesTitle) Then
        Err.Raise 384
    End If
    mudtProps.SeriesTitle = val
    SetSeriesTitle = ValidateObject("SERIESTITLE")
    mudtProps.IsDirty = True
End Function

Public Property Get Author() As String
    Author = Trim(mudtProps.Author)
End Property
Public Function SetAuthor(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetAuthor = True
    If Len(val) > Len(mudtProps.Author) Then
        Err.Raise 384
    End If
    mudtProps.Author = val
    SetAuthor = ValidateObject("AUTHOR")
    mudtProps.IsDirty = True
End Function

Public Property Get Publisher() As String
    Publisher = Trim(mudtProps.Publisher)
End Property
Public Function SetPublisher(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPublisher = True
    If Len(val) > Len(mudtProps.Publisher) Then
        Err.Raise 384
    End If
    mudtProps.Publisher = val
    SetPublisher = ValidateObject("PUBLISHER")
    mudtProps.IsDirty = True
End Function

Public Property Get SupplierID() As String
    SupplierID = mudtProps.SupplierID
End Property
Public Property Get DealID() As String
    DealID = mudtProps.DealID
End Property
Public Property Get SpecialVat() As Boolean
    SpecialVat = mudtProps.SpecialVat
End Property
Public Property Let SpecialVat(val As Boolean)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SpecialVat = val
    SetDirty True
End Property
Public Property Get IsServiceItem() As Boolean
    IsServiceItem = (FNS(mudtProps.ProductType) = "N")
End Property
Public Function SetServiceItem()
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "N"
    SetDirty True
End Function
Public Property Get IsGeneralProduct() As Boolean
    IsGeneralProduct = (FNS(mudtProps.ProductType) = "G")
End Property
Public Function SetGeneralProduct()
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "G"
    SetDirty True
End Function
Public Property Get IsMagsEtc() As Boolean
    IsMagsEtc = (FNS(mudtProps.ProductType) = "M")
End Property
Public Function SetMagsEtc()
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "M"
    SetDirty True
End Function
Public Property Get IsBook() As Boolean
    IsBook = (FNS(mudtProps.ProductType) = "B")
End Property
Public Function SetBook()
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "B"
    SetDirty True
End Function

Public Property Get ProductType() As String
    ProductType = FNS(mudtProps.ProductType)
End Property
Public Property Get Obsolete() As Boolean
    Obsolete = mudtProps.Obsolete
End Property
Public Property Let Obsolete(val As Boolean)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Obsolete = val
    SetDirty True
End Property
Public Property Get VATRateToUse() As Double
    If SpecialVat Then
        VATRateToUse = VATRate
    Else
        VATRateToUse = oPC.Configuration.VATRate
    End If
End Property
Public Property Get VATRateToUseF() As String
    If SpecialVat Then
        VATRateToUseF = VATRateF
    Else
        VATRateToUseF = oPC.Configuration.VATRateF
    End If
End Property
Public Property Get VATRate() As Double
    VATRate = mudtProps.VATRate
End Property
Public Property Get VATRateF() As String
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
End Property
Public Property Let VATRate(val As Double)
    If mcolStack.Count = 0 Then Err.Raise 383
    If val <> oPC.Configuration.VATRate Then
        Me.SpecialVat = True
        mudtProps.VATRate = val
        SetDirty True
    End If
End Property
Public Function SetVAT(val As String) As Boolean
Dim bOK As Boolean
    
    bOK = SetField_DOUBLE(mudtProps.VATRate, val, "VATRATE")
    If bOK Then
        bOK = ValidateObject("VATRATE")
    End If
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        mudtProps.SpecialVat = True
    End If
    SetVAT = bOK

End Function
'Public Property Let Availability(Val As String)
'    If Not flgEditing Then Err.Raise 383
'    If IsNull(Val) Then GoTo EXIT_Handler
'    If Not IsNull(Val) Then mudtProps.Availability = Val
'    flgDirty = True
'EXIT_Handler:
'    Exit Property
'End Property
'Public Property Get Availability() As String
'    Availability = mudtProps.Availability
'End Property
'Public Property Get QtyOnHand() As Long
'    QtyOnHand = mudtProps.QtyOnHand
'End Property
'
'Public Property Let QtyOnHand(Val As Long)
'    If Not flgEditing Then Err.Raise 383
'    If IsNull(Val) Then GoTo EXIT_Handler
'    If Not IsNull(Val) Then mudtProps.StockBalance = Val
'    flgDirty = True
'EXIT_Handler:
'    Exit Property
'End Property
Public Property Let LastCopySerial(val As Integer)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LastCopySerial = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
End Property
Public Property Get LastCopySerial() As Integer
    LastCopySerial = mudtProps.LastCopySerial
End Property

Public Property Let DealID(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsNull(val) Then GoTo EXIT_Handler
    mudtProps.DealID = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
End Property

Public Property Get Categories() As z_TextList
    Set Categories = tlCategories
End Property
Public Property Get CategoryHeadings() As z_TextList
    Set CategoryHeadings = tlCatalogueHeadings
End Property

Public Sub OutOnApproPerTP(ByVal pTPID As Long, ByVal pPID As String, ByRef iQty As Integer, ByRef lngApproID, ByRef lngNumOfApproLines)
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL

    On Error GoTo ERR_Handler
    
    Set rs = New ADODB.Recordset
    GoTo EXIT_Handler
    MsgBox "Skipping"
    Set oBatch = New z_SQL
    lngResult = oBatch.RunGetRecordset("q_ApprosOutstandingPerTitle", enQuery, Array(pPID, pTPID), "", rs)
    If rs.eof Then GoTo EXIT_Handler
    
    Do While Not rs.eof
        If lngNumOfApproLines > 1 Then GoTo EXIT_Handler
        iQty = rs!bal
        lngApproID = rs!APPL_APP_ID
        lngNumOfApproLines = lngNumOfApproLines + 1
        rs.MoveNext
    Loop
    
EXIT_Handler:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub
Private Function ValidateObject(pFld As String) As Boolean
Dim oProdCode  As New z_ProdCode
Dim bValid As Boolean
    bValid = True
    If Not (Len(Trim(mudtProps.Title)) > 3 Or Len(Trim(mudtProps.Title)) = 0) Then
        mobjValid.BreakRule "Title", True
        If pFld = "TITLE" Then bValid = False
    Else
        mobjValid.BreakRule "TITLE", False
    End If
    If Not (Len(Trim(mudtProps.SubTitle)) > 3 Or Len(Trim(mudtProps.SubTitle)) = 0) Then
        mobjValid.BreakRule "subTitle", True
        If pFld = "SUBTITLE" Then bValid = False
    Else
        mobjValid.BreakRule "SUBTITLE", False
    End If
    oProdCode.LoadNew mudtProps.code, mudtProps.EAN, mudtProps.ProductType = "B", "", True
    If oProdCode.CodesOK Then
        mobjValid.BreakRule "CODE", False
    Else
        mobjValid.BreakRule "Code", True
    End If
    If Not (Len(Trim(mudtProps.Author)) > 2 Or Len(Trim(mudtProps.Author)) = 0) Then
        mobjValid.BreakRule "Author", True
        If pFld = "AUTHOR" Then bValid = False
    Else
        mobjValid.BreakRule "AUTHOR", False
    End If
    If Not (Len(Trim(mudtProps.Publisher)) > 2 Or Len(Trim(mudtProps.Publisher)) = 0) Then
        mobjValid.BreakRule "Publisher", True
        If pFld = "PUBLISHER" Then bValid = False
    Else
        mobjValid.BreakRule "PUBLISHER", False
    End If
    If Not (Len(Trim(mudtProps.SeriesTitle)) > 2 Or Len(Trim(mudtProps.SeriesTitle)) = 0) Then
        mobjValid.BreakRule "SeriesTitle", True
        If pFld = "SERIESTITLE" Then bValid = False
    Else
        mobjValid.BreakRule "SERIESTITLE", False
    End If
'    If Len(Trim(mudtProps.EAN)) > 0 Then oProdCode.Load Trim$(mudtProps.EAN)
'    If (oProdCode.IsEAN Or Trim$(mudtProps.EAN) = "") Then
'        mobjValid.BreakRule "EAN", False
'    Else
'        mobjValid.BreakRule "EAN", True
'        If pFld = "EAN" Then bValid = False
'    End If
    If (mudtProps.RRP >= 0) Then
        mobjValid.BreakRule "RRP", False
    Else
        mobjValid.BreakRule "RRP", True
        If pFld = "RRP" Then bValid = False
    End If
    If (mudtProps.Cost >= 0) Then
        mobjValid.BreakRule "COST", False
    Else
        mobjValid.BreakRule "COST", True
        If pFld = "COST" Then bValid = False
    End If
    If (mudtProps.SP >= 0) Then
        mobjValid.BreakRule "SP", False
    Else
        mobjValid.BreakRule "SP", True
        If pFld = "SP" Then bValid = False
    End If
    If (mudtProps.UKPrice >= 0) Then
        mobjValid.BreakRule "UKPRICE", False
    Else
        mobjValid.BreakRule "UKPRICE", True
        If pFld = "UKPRICE" Then bValid = False
    End If

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
            strAllRules = mcolClassors(Left(pRawors, iMarker - 1))
        Else
            strAllRules = mcolClassors(pRawors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawors, ",")
            If iMarker > 0 Then
                strRule = mcolClassors(Mid(pRawors, iStart, iMarker - iStart))
            Else
                strRule = mcolClassors(Mid(pRawors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
End Function
Private Sub LoadClassorsCollection()
    Set mcolClassors = New Collection
    mcolClassors.Add "Name too short", "NAME"
    mcolClassors.Add "Title too short", "TITLE"
    mcolClassors.Add "Author too short", "AUTHOR"
    mcolClassors.Add "Subtitle too short", "SUBTITLE"
    mcolClassors.Add "Publisher too short", "PUBLISHER"
    mcolClassors.Add "Publication place too short", "PUBPLACE"
    mcolClassors.Add "Publication date too short", "PUBDATE"
    mcolClassors.Add "Edition too short", "EDITION"
    mcolClassors.Add "Invalid EAN", "EAN"
    mcolClassors.Add "Invalid code", "CODE"
    mcolClassors.Add "Note too short", "NOTE"
    mcolClassors.Add "Invalid cost", "COST"
    mcolClassors.Add "Invalid R.R.P.", "RRP"
    mcolClassors.Add "Invalid S.P.", "SP"
End Sub

Private Sub SetDirty(pVal As Boolean)
    mudtProps.IsDirty = pVal
 '   RaiseEvent Dirty(pVal)
End Sub

Public Property Get Stock6() As Integer
    Stock6 = mudtProps.StckAgeQty6Mnths
End Property
Public Property Get Stock12() As Integer
    Stock12 = mudtProps.StckAgeQty12Mnths
End Property
Public Property Get Stock18() As Integer
    Stock18 = mudtProps.StckAgeQty18Mnths
End Property
Public Property Get Stock18Plus() As Integer
    Stock18Plus = mudtProps.StckAgeQty18MnthsPlus
End Property
Public Property Get StockAgedDate() As Integer
    StockAgedDate = mudtProps.StckAgeDate
End Property
Public Property Get DateRecordAddedF() As String
    DateRecordAddedF = mudtProps.DateAdded
End Property
Public Property Get DateLastModifiedF() As String
    DateLastModifiedF = Format(mudtProps.DateLastModified, "dd/mm/yyyy hh:nn")
End Property

Private Function SetField_CURRENCY(fld As Currency, val As String, pValidationName As String)
Dim cTemp As Currency
Dim bTemp As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_CURRENCY = True
    If Trim$(val) = "" Then
        cTemp = 0
    ElseIf Not ConvertToCurr(val, cTemp) Then
        SetField_CURRENCY = False
        Exit Function
    End If
    fld = cTemp
End Function
Private Function SetField_LONG(fld As Long, val As String, pValidationName As String)
Dim lngTemp As Long
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_LONG = True
    If Trim$(val) = "" Then
        lngTemp = 0
    ElseIf Not ConvertToLng(val, lngTemp) Then
        SetField_LONG = False
        Exit Function
    End If
    fld = lngTemp
End Function
Private Function SetField_DATE(fld As Date, val As String, pValidationName As String)
Dim dteTemp As Date
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DATE = True
    If Trim$(val) = "" Then
        dteTemp = 0
    ElseIf Not ConvertToDate(val, dteTemp) Then
        SetField_DATE = False
        Exit Function
    End If
    fld = dteTemp
End Function
Private Function SetField_STRING(fld As String, val As String, pValidationName As String)
Dim strTemp As String

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_STRING = True
    strTemp = val
    If Len(strTemp) > Len(fld) Then
        Err.Raise vbObjectError + 1001, "String value too long"
    End If
    fld = strTemp
End Function
Private Function SetField_DOUBLE(fld As Double, val As String, pValidationName As String)
Dim dblTEMP As Double
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DOUBLE = True
    If Trim$(val) = "" Then
        dblTEMP = 0
    ElseIf Not ConvertToDBL(val, dblTEMP) Then
        SetField_DOUBLE = False
        Exit Function
    End If
    fld = dblTEMP
End Function


