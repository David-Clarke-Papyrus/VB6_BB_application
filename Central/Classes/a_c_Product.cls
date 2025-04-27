VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Product"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Copies"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Copy"
Option Explicit

'Product class for Application: CENTRAL

Private mudtProps As ProductProps
Private mudtSave As ProductProps
Private mcolStores As ch_StoreP
Private mcolStack As Collection

Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private WithEvents oProdCode As z_ProdCode
Attribute oProdCode.VB_VarHelpID = -1
Private mcolClassErrors As Collection
Event Valid(msg As String)
Event CodeToBeGenerated()
Event RedisplayCodes()
Dim oBF As a_BookFind

Private flgEditing As Boolean
Private Function GetState() As String
    On Error GoTo ErrHandler
  Dim udtData As ProductData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo ErrHandler
  Dim udtData As ProductData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetState(buffer)", buffer
End Sub

Public Function GetSuperState() As String
    On Error GoTo ErrHandler
Dim objPB As PropertyBag
  
    Set objPB = New PropertyBag
    objPB.WriteProperty "State", GetState
    objPB.WriteProperty "Stores", mcolStores.GetSuperState
    GetSuperState = objPB.Contents
    Set objPB = Nothing
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.GetSuperState"
End Function

Public Sub SetSuperState(buffer As String)
    On Error GoTo ErrHandler
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim strBuff As String
    Set objPB = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    SetState objPB.ReadProperty("State")
        mcolStores.SetSuperState objPB.ReadProperty("Stores")
    Set objPB = Nothing
    mobjValid.RuleBroken "TITLE", False
    mobjValid.RuleBroken "AUTHOR", False
    mobjValid.RuleBroken "CODE", False

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSuperState(buffer)", buffer
End Sub


'Public Function Exists(pCODE As String) As Boolean
'    On Error GoTo errHandler
'Dim objPersist As a_C_Product_p
'    Set objPersist = New a_C_Product_p
'    Exists = objPersist.Exists(pCODE)
'    Set objPersist = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.Exists(pCODE)", pCODE
'End Function
'Public Function ExistsEx(pCODE As String, PID As String) As Boolean
'    On Error GoTo errHandler
'Dim objPersist As a_C_Product_p
'    Set objPersist = New a_C_Product_p
'    ExistsEx = objPersist.Exists(pCODE, PID)
'    Set objPersist = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.ExistsEx(pCODE,PID)", Array(pCODE, PID)
'End Function

Public Property Get Stores() As ch_StoreP
    On Error GoTo ErrHandler
    Set Stores = mcolStores
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stores"
End Property

Public Property Get IsEditing() As Boolean
    On Error GoTo ErrHandler
    IsEditing = mcolStack.Count > 0
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsEditing"
End Property
Public Sub BeginEdit()
    On Error GoTo ErrHandler
    mcolStores.BeginEdit
    mcolStack.Add GetState
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo ErrHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
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
    ErrorIn "a_Product.CancelEdit"
End Sub
'Private Sub Checkcodes()
'Dim oPC As z_ProdCode
'    Set oPC = New z_ProdCode
'    oPC.Load Me.Code, True
'
'End Sub
Public Sub ApplyEdit(Optional plResult As Long, Optional pMsg As String)
    On Error GoTo ErrHandler
Dim objPersist As a_C_Product_p
Dim oProdCode As z_ProdCode
Dim strTemp As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    plResult = 0
    If mcolStack.Count = 0 Then Err.Raise 445
    
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        mudtProps.SpecialVat = True
    End If
    oPC.COShort.CommandTimeout = 7
    oPC.COShort.BeginTrans
    mcolStores.ApplyEdit
    Set objPersist = New a_C_Product_p
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject Trim$(mudtProps.ID)
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    strTemp = objPersist.Save(GetSuperState)
    If UCase(Left(strTemp, 7)) <> "TIMEOUT" And UCase(Left(strTemp, 9)) <> "DUPLICATE" Then
        mcolStack.Remove mcolStack.Count
        SetSuperState strTemp
        mudtProps.IsNew = False
    Else 'save failed
        mcolStores.BeginEdit
        If Not IsMissing(pMsg) Then
            pMsg = strTemp
        End If
        plResult = 99
        oPC.COShort.RollbackTrans
        Set objPersist = Nothing
        GoTo EXIT_Handler
    End If
  Else
    mcolStack.Remove mcolStack.Count
  End If
  Set objPersist = Nothing
  mudtProps.IsDirty = False
  oPC.COShort.CommitTrans
  oPC.COShort.CommandTimeout = 0
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
'ERR_Handler:
'    plResult = 101
'    MsgBox "Unknown Error" & Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
ErrHandler:
    ErrPreserve

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ApplyEdit(plResult,pMsg)", Array(plResult, pMsg), EA_ROLLBACK + EA_RERAISE, oPC.COShort
End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Set mcolStack = New Collection
    Set mcolStores = New ch_StoreP
    Set mobjValid = New z_BrokenRules
    Set mcolClassErrors = New Collection
    Set oProdCode = New z_ProdCode
    LoadClassErrorsCollection
    mobjValid.RuleBroken "TITLE", True
    mudtProps.IsNew = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Class_Initialize"
End Sub
Public Sub Initialize(pbKeepsCopies As Boolean)
    On Error GoTo ErrHandler
    mudtProps.SpecialVat = False
    mudtProps.VATRate = oPC.Configuration.VATRate
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Initialize(pbKeepsCopies)", pbKeepsCopies
End Sub
Private Sub Class_Terminate()
    On Error GoTo ErrHandler
    If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    
    Set mcolStack = Nothing
    Set mcolStores = Nothing
    Set mobjValid = Nothing
    Set mcolClassErrors = Nothing
    Set oProdCode = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo ErrHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsValid"
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo ErrHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsDirty"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo ErrHandler
  IsNew = mudtProps.IsNew
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsNew"
End Property
Public Function GetSTatus()
    On Error GoTo ErrHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.GetSTatus"
End Function

Private Sub mobjValid_BrokenRule(errors As String)
    On Error GoTo ErrHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_BrokenRule(errors)", errors
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    On Error GoTo ErrHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_RuleUnbroken(errors)", errors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo ErrHandler
    RaiseEvent Valid("")
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo ErrHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
   ' RaiseEvent Valid("")
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_Status(pMsg)", pMsg
End Sub
Public Sub PrepareForReload()
    On Error GoTo ErrHandler
    mudtProps.IsNew = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PrepareForReload"
End Sub
Public Function Load(PID As String, pCatID As Long, Optional pCode As String, Optional strTime As String) As Long  ', Optional pbReload As Boolean
    On Error GoTo ErrHandler
'Returns 0 if found on db or loaded and saved from Bookfind
'Returns 99 if an error occurs on saving from Bookfind
'Returns 98 if not found on db or Bookfind
Dim strCode As String
Dim lngResult As Long
Dim lngFound As Long
Dim lngPK As Long
Dim iProdID As Long
Dim tmpBuffer As String
Dim objPersist As a_C_Product_p
Dim i As Integer
Dim tmp As String
Dim iPos As Integer


    strCode = FNS(pCode)

    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445

    Set objPersist = New a_C_Product_p
start:
'option 1
    lngFound = 99
    If Len(PID) > 1 Then
        tmp = objPersist.Fetch(lngResult, PID) 'returns 0 if found; 99 if not found
        If tmp > "" Then
            lngFound = 0
            SetSuperState tmp
            GoTo EXIT_Handler
        Else
            lngFound = 0
            GoTo EXIT_Handler
        End If
    End If
'option 2
    If IsISBN13(strCode) Then
        tmpBuffer = objPersist.Fetch(lngResult, , , strCode)
        If tmpBuffer > "" Then
            SetSuperState tmpBuffer
            lngFound = 0
            GoTo EXIT_Handler
        Else
            lngFound = 99
        End If
        'Convert ISBN13 to ISBN10 if possible so that next search may work
        strCode = ISBN13to10(strCode)
    End If
'option 3
    If IsISBN10(strCode) Or IsHashCode(strCode) Or IsPrivateCode(strCode) Then   'pCode s not an ISBN-13 or an EAN
        tmpBuffer = objPersist.Fetch(lngResult, , strCode)
        If tmpBuffer > "" Then
            SetSuperState tmpBuffer
        Else
            lngFound = 99
        End If
    End If

'we only get here if we have not found the record
    If IsISBN13(FNS(pCode)) Or IsISBN10(FNS(pCode)) Then
        i = 1
            Do While oPC.Configuration.LookupSequence(i) > "" And lngResult <> 0
             Select Case oPC.Configuration.LookupSequence(i)
             Case "BF"
                    Set oBF = New a_BookFind
                    lngResult = oBF.FetchFromBF(FNS(pCode))
                    If lngResult = 0 Then  'Found a record
                        lngFound = 0
                        LoadProductFromBF
                        SetSuperState objPersist.Save(GetSuperState)
                        PID = mudtProps.ID
                        Set oBF = Nothing
                        GoTo start
                        mudtProps.IsDirty = True
                        mudtProps.IsNew = False
                    Else
                        PID = 0 'product not on db and not on Bookfind
                        lngFound = 99
                    End If
                    Set oBF = Nothing
             Case "WH"
             End Select
            i = i + 1
        Loop
    End If
   ' Set oProdCode = Nothing
    Load = lngFound

    
EXIT_Handler:
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Load(PID,pCatID,pCODE,strTime)", Array(PID, pCatID, pCode, strTime)
End Function


Public Sub Delete()
    On Error GoTo ErrHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Delete"
End Sub

Public Function LookOnBookfind(pCode As String) As Integer
    On Error GoTo ErrHandler
Dim iresult As Long
    Set oBF = New a_BookFind
    iresult = oBF.DestroyResultsList
    iresult = oBF.FetchFromBF(pCode)
    If iresult > 0 Then
        LoadProductFromBF
    End If
    LookOnBookfind = iresult
    Set oBF = Nothing
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LookOnBookfind(pCODE)", pCode
End Function
Private Sub LoadProductFromBF()
    On Error GoTo ErrHandler
    mudtProps.Code = oBF.Code
    mudtProps.Author = oBF.MainAuthor
    mudtProps.Title = oBF.Title
    mudtProps.SubTitle = oBF.SubTitle
    mudtProps.Edition = oBF.Edition
    mudtProps.Publisher = oBF.PublisherName
    mudtProps.Note = oBF.Note
    mudtProps.SeriesTitle = oBF.SeriesTitle
    mudtProps.Availability = Left(oBF.Availability, 1)
    mudtProps.UKPrice = oBF.UKPrice
    mudtProps.USPrice = oBF.USPrice
    mudtProps.RRP = oBF.LocalPrice * 100 'oPC.Configuration.DefaultCurrencyDivisor
    mudtProps.PublicationDate = oBF.PublicationDate
    mudtProps.Description = oBF.Description
    mudtProps.MainSupplierName = oBF.MainSupplierName
    mudtProps.BindingCode = oBF.BindingCode
    mudtProps.BFClassification = oBF.BFClassification
    mudtProps.ProductType = "B"
  '  MsgBox "Possibly problem assigning properties directly -specially prices"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoadProductFromBF"
End Sub
Public Property Get PublicationDate() As String
    On Error GoTo ErrHandler
    PublicationDate = Trim(mudtProps.PublicationDate)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PublicationDate"
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
    Exit Function
End Function
Public Property Get PublicationPlace() As String
    On Error GoTo ErrHandler
    PublicationPlace = Trim(mudtProps.PublicationPlace)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PublicationPlace"
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
    Exit Function
End Function
Public Property Get Edition() As String
    On Error GoTo ErrHandler
    Edition = Trim(mudtProps.Edition)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Edition"
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
    Exit Function
End Function

Public Property Get PID() As String
    On Error GoTo ErrHandler
    PID = FNS(mudtProps.ID)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PID"
End Property
Public Property Get DefaultDeliveryDays() As Integer
    On Error GoTo ErrHandler
    DefaultDeliveryDays = IIf(mudtProps.DefaultDeliveryDays = 0, 30, mudtProps.DefaultDeliveryDays)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DefaultDeliveryDays"
End Property

Public Property Get Description() As String
    On Error GoTo ErrHandler
    Description = Trim$(mudtProps.Description)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Description"
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
    Exit Function
End Function
Public Property Get BFDistributorCode() As String
    On Error GoTo ErrHandler
    BFDistributorCode = Trim$(mudtProps.BFDistributorCode)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BFDistributorCode"
End Property
Public Function SetBFDistributorCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetBFDistributorCode = True
    If Len(val) > Len(mudtProps.BFDistributorCode) Then
        Err.Raise 384
    End If
    mudtProps.BFDistributorCode = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Note() As String
    On Error GoTo ErrHandler
    Note = Trim$(mudtProps.Note)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Note"
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
    Exit Function
End Function
Public Property Get Code() As String
    On Error GoTo ErrHandler
    Code = FNS(mudtProps.Code)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Code"
End Property
Public Property Get CodeF() As String
    On Error GoTo ErrHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CodeF"
End Property
Public Property Get BottomOfDocument() As Boolean
    On Error GoTo ErrHandler
    BottomOfDocument = (Right(mudtProps.Code, 2) = "_B")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BottomOfDocument"
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
Public Function SetCode(val As String, Optional pIsISBN As Boolean) As Boolean
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetCode = True
'    If Len(val) > Len(mudtProps.Code) Then
'        Err.Raise 384
'    End If
    mudtProps.Code = val
'    oProdCode.Load val, True
'    If Not IsMissing(pIsISBN) Then pIsISBN = oProdCode.IsISBN
    SetCode = ValidateObject("CODE")
'    If EAN = "" And oProdCode.EANExists Then
'        SetEAN oProdCode.EAN
'    End If
    mudtProps.IsDirty = True
 '   If val = "#" Then RaiseEvent CodeToBeGenerated
    RaiseEvent RedisplayCodes
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCode(val,pIsISBN)", Array(val, pIsISBN)
End Function
Public Property Get EAN() As String
    On Error GoTo ErrHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.EAN"
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
    Exit Function
End Function
Public Property Get TitleAuthorPublisher() As String
    On Error GoTo ErrHandler
Dim tmp As String

    tmp = IIf(mudtProps.Status = "O", "***OP***", IIf(mudtProps.Status = "R", "***RP***", "")) & FNS(mudtProps.Title)
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Author)) > 0, "(" & Trim$(mudtProps.Author) & ")", "")
    End If
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Publisher)) > 0, "(" & Trim$(mudtProps.Publisher) & ")", "")
    End If
    TitleAuthorPublisher = tmp
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TitleAuthorPublisher"
End Property
Public Property Get TitleAuthorPublisherL(pLen As Integer) As String
    On Error GoTo ErrHandler
'Dim tmp As String
'
'    tmp = Trim$(mudtProps.Title)
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(Trim$(mudtProps.Author)) > 0, "(" & Trim$(mudtProps.Author) & ")", "")
'    End If
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(Trim$(mudtProps.Publisher)) > 0, "(" & Trim$(mudtProps.Publisher) & ")", "")
'    End If
    TitleAuthorPublisherL = Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", "")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TitleAuthorPublisherL(pLen)", pLen
End Property
Public Property Get BICDescription() As String
    On Error GoTo ErrHandler
    BICDescription = FNS(mudtProps.BICDescription)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BICDescription"
End Property
Public Property Get BIC() As String
    On Error GoTo ErrHandler
    BIC = Trim$(mudtProps.BIC)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BIC"
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
    Exit Function
End Function

Public Property Get Section() As String
    On Error GoTo ErrHandler
    Section = FNS(mudtProps.Section)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Section"
End Property
'Public Function SetSection(val As String) As Boolean
'    On Error GoTo errHandler
'
'    If oPC.Configuration.Sections.ValidateNewSetMember(mudtProps.Section, val) Then
'        SetSection = True
'        SetSection = ValidateObject("SECTION")
'        mudtProps.IsDirty = True
'    End If
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SetSection(val)", val
'End Function
'Public Function SetSectionALL(val As String) As Boolean
'    On Error GoTo errHandler
'Dim ar() As String
'Dim i As Integer
'Dim j As Integer
'Dim iCount As Integer
'Dim tmp As String
'
'    If val = "" And oPC.Configuration.EnforceSections = False Then
'        SetSectionALL = True
'        If FNS(mudtProps.Section) <> val Then
'            mudtProps.Section = val
'            mudtProps.IsDirty = True
'        End If
'        Exit Function
'    End If
'
'    If oPC.Configuration.Sections.ValidateEntireSet(val) Then
'        mudtProps.Section = val
'        SetSectionALL = True
'        SetSectionALL = ValidateObject("SECTION")
'        mudtProps.IsDirty = True
'    End If
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SetSectionALL(val)", val
'End Function
Public Property Get Comment() As String
    On Error GoTo ErrHandler
    Comment = Trim$(mudtProps.Comment)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Comment"
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
    Exit Function
End Function

Public Property Get Summary() As String
    On Error GoTo ErrHandler
    Summary = Trim$(mudtProps.Summary)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Summary"
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
    Exit Function
End Function


Public Property Get BindingCode() As String
    On Error GoTo ErrHandler
    BindingCode = Trim$(mudtProps.BindingCode)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BindingCode"
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
    Exit Function
End Function


Public Property Get Title() As String
    On Error GoTo ErrHandler
    Title = FNS(mudtProps.Article) & FNS(mudtProps.Title)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Title"
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
    Exit Function
End Function
Public Property Get SubTitle() As String
    On Error GoTo ErrHandler
    SubTitle = Trim(mudtProps.SubTitle)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SubTitle"
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
    Exit Function
End Function
Public Property Get FlagText() As String
    On Error GoTo ErrHandler
    FlagText = Trim(mudtProps.FlagText)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.FlagText"
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
    Exit Function
End Function

Public Property Get CategoryID() As Long
    On Error GoTo ErrHandler
    CategoryID = mudtProps.CategoryID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CategoryID"
End Property
Public Property Get ProductTypeID() As Long
    On Error GoTo ErrHandler
    ProductTypeID = mudtProps.ProductTypeID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ProductTypeID"
End Property
'Public Property Get CategoryName() As String
'    CategoryName = tlSections.Item(mudtProps.CategoryID)
'End Property
Public Function SetCategoryID(val As Long) As Boolean
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCategoryID = True
    mudtProps.CategoryID = val
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCategoryID(val)", val
End Function
'Public Property Get ProductTypeName() As String
'    ProductTypeName = tlProductTypes.Item(mudtProps.ProductTypeID)
'End Property
Public Function SetProductTypeID(val As Long) As Boolean
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetProductTypeID = True
    mudtProps.ProductTypeID = val
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetProductTypeID(val)", val
End Function
Public Property Get CatalogueheadingID() As Long
    On Error GoTo ErrHandler
    CatalogueheadingID = mudtProps.CatalogueheadingID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CatalogueheadingID"
End Property
Public Function SetCatalogueheadingID(val As Long) As Boolean
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCatalogueheadingID = True
    mudtProps.CatalogueheadingID = val
 '   SetCategoryID = ValidateObject(Val, "CATEGORYID")
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCatalogueheadingID(val)", val
End Function
'Public Property Get CatalogueHeadingName() As String
'    CatalogueHeadingName = tlCatalogueHeadings.Item(mudtProps.CatalogueheadingID)
'End Property
'
Public Property Get BFClassification() As String
    On Error GoTo ErrHandler
    BFClassification = mudtProps.BFClassification
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BFClassification"
End Property

Public Property Get PriceLastDelivered() As Long
    On Error GoTo ErrHandler
    PriceLastDelivered = mudtProps.PriceLastDelivered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastDelivered"
End Property
Public Property Get PriceLastDeliveredF() As String
    On Error GoTo ErrHandler
    If mudtProps.PriceLastDelivered = 0 Then
        PriceLastDeliveredF = ""
    Else
        PriceLastDeliveredF = Format(mudtProps.PriceLastDelivered / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastDeliveredF"
End Property
Public Function SetPriceLastDelivered(val As String) As Boolean
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetPriceLastDelivered(val)", val
End Function
Public Property Get PriceLastCountedF() As String
    On Error GoTo ErrHandler
    If mudtProps.PriceLastCounted = 0 Then
        PriceLastCountedF = ""
    Else
        PriceLastCountedF = Format(mudtProps.PriceLastCounted / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastCountedF"
End Property
Public Function SetPriceLastCounted(val As String) As Boolean
    On Error GoTo ErrHandler
Dim lngPriceLastCounted As Integer
    SetPriceLastCounted = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngPriceLastCounted = 0
    ElseIf (Not ConvertToInt(val, lngPriceLastCounted)) Then
        SetPriceLastCounted = False
        Exit Function
    End If
    mudtProps.PriceLastDelivered = lngPriceLastCounted
    SetPriceLastCounted = ValidateObject("PriceLastCOunted")
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetPriceLastCounted(val)", val
End Function

Public Property Get PricelastSold() As Long
    On Error GoTo ErrHandler
    PricelastSold = mudtProps.PricelastSold
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PricelastSold"
End Property
Public Property Get PriceLastSoldF() As String
    On Error GoTo ErrHandler
    If mudtProps.PricelastSold = 0 Then
        PriceLastSoldF = ""
    Else
        PriceLastSoldF = Format(mudtProps.PricelastSold / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    End If

    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastSoldF"
End Property


Public Property Get PriceLastOrdered() As Long
    On Error GoTo ErrHandler
    PriceLastOrdered = mudtProps.PriceLastOrdered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastOrdered"
End Property
Public Property Get PriceLastOrderedF() As String
    On Error GoTo ErrHandler
    If mudtProps.PriceLastOrdered = 0 Then
        PriceLastOrderedF = ""
    Else
        PriceLastOrderedF = Format(mudtProps.PriceLastOrdered / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    End If

    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastOrderedF"
End Property
Public Function SetPriceLastOrdered(val As String) As Boolean
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetPriceLastOrdered(val)", val
End Function

Public Property Let DateLastDelivered(val As Date)
    On Error GoTo ErrHandler
    mudtProps.DateLastDelivered = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastDelivered(val)", val
End Property
Public Property Get DateLastDelivered() As Date
    On Error GoTo ErrHandler
    DateLastDelivered = mudtProps.DateLastDelivered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastDelivered"
End Property
Public Property Get DateLastDeliveredF() As String
    On Error GoTo ErrHandler
    If mudtProps.DateLastDelivered = CDate(0) Then
        DateLastDeliveredF = ""
    Else
        DateLastDeliveredF = Format(mudtProps.DateLastDelivered, "dd/mm/yyyy")
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastDeliveredF"
End Property

Public Function SetDateLastDelivered(val As String) As Boolean
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetDateLastDelivered(val)", val
End Function
Public Property Let DateLastOrdered(val As Date)
    On Error GoTo ErrHandler
    mudtProps.DateLastOrdered = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastOrdered(val)", val
End Property
Public Property Get DateLastOrdered() As Date
    On Error GoTo ErrHandler
    DateLastOrdered = mudtProps.DateLastOrdered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastOrdered"
End Property
Public Property Get DateLastOrderedF() As String
    On Error GoTo ErrHandler
    If mudtProps.DateLastOrdered = CDate(0) Then
        DateLastOrderedF = ""
    Else
        DateLastOrderedF = Format(mudtProps.DateLastOrdered, "dd/mm/yyyy")
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastOrderedF"
End Property
Public Property Get DateLastSold() As Date
    On Error GoTo ErrHandler
    DateLastSold = mudtProps.DateLastSold
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastSold"
End Property
Public Property Get DateLastSoldF() As String
    On Error GoTo ErrHandler
    If mudtProps.DateLastSold = CDate(0) Then
        DateLastSoldF = ""
    Else
        DateLastSoldF = Format(mudtProps.DateLastSold, "dd/mm/yyyy")
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastSoldF"
End Property

Public Function SetDateLastOrdered(val As String) As Boolean
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetDateLastOrdered(val)", val
End Function


Public Property Let DateLastCounted(val As Date)
    On Error GoTo ErrHandler
    mudtProps.DateLastCounted = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastCounted(val)", val
End Property
Public Property Get DateLastCounted() As Date
    On Error GoTo ErrHandler
    DateLastCounted = mudtProps.DateLastCounted
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastCounted"
End Property
Public Property Get dateLastCountedF() As String
    On Error GoTo ErrHandler
    If mudtProps.DateLastCounted = CDate(0) Then
        dateLastCountedF = ""
    Else
        dateLastCountedF = Format(mudtProps.DateLastCounted, "dd/mm/yyyy")
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.dateLastCountedF"
End Property

Public Function SetdateLastCounted(val As String) As Boolean
    On Error GoTo ErrHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetdateLastCounted = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetdateLastCounted = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateLastCounted", False
    Else
        mobjValid.RuleBroken "DateLastCounted", True
    End If
    mudtProps.DateLastCounted = dteDate
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetdateLastCounted(val)", val
End Function





Public Property Get QtyLastCounted() As Long
    On Error GoTo ErrHandler
    QtyLastCounted = mudtProps.QtyLastCounted
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastCounted"
End Property
Public Property Get QtyLastCountedF() As String
    On Error GoTo ErrHandler
    QtyLastCountedF = mudtProps.QtyLastCounted
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastCountedF"
End Property
Friend Property Let QtyLastCounted(val As Long)
    On Error GoTo ErrHandler
    mudtProps.QtyLastCounted = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastCounted(val)", val
End Property

Public Property Get QtyLastDelivered() As Long
    On Error GoTo ErrHandler
    QtyLastDelivered = mudtProps.QtyLastDelivered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastDelivered"
End Property
Public Property Get QtyLastDeliveredF() As String
    On Error GoTo ErrHandler
    QtyLastDeliveredF = mudtProps.QtyLastDelivered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastDeliveredF"
End Property
Friend Property Let QtyLastDelivered(val As Long)
    On Error GoTo ErrHandler
    mudtProps.QtyLastDelivered = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastDelivered(val)", val
End Property
Public Property Get QtyCopiesOnHand() As Long
    On Error GoTo ErrHandler
    QtyCopiesOnHand = mudtProps.QtyCopiesOnHand
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyCopiesOnHand"
End Property
Public Property Get QtyCopiesOnHandF() As String
    On Error GoTo ErrHandler
    QtyCopiesOnHandF = mudtProps.QtyCopiesOnHand
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyCopiesOnHandF"
End Property
Friend Property Let QtyCopiesOnHand(val As Long)
    On Error GoTo ErrHandler
    mudtProps.QtyCopiesOnHand = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyCopiesOnHand(val)", val
End Property
Public Property Get QtyOnHand() As Long
    On Error GoTo ErrHandler
    QtyOnHand = FNN(mudtProps.QtyOnHand)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnHand"
End Property
Public Property Get QtyOnHandF() As String
    On Error GoTo ErrHandler
    QtyOnHandF = mudtProps.QtyOnHand
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnHandF"
End Property
Friend Property Let QtyOnHand(val As Long)
    On Error GoTo ErrHandler
    mudtProps.QtyOnHand = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnHand(val)", val
End Property
Public Property Get QtyTotalSold() As Long
    On Error GoTo ErrHandler
    QtyTotalSold = mudtProps.QtyTotalSold
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyTotalSold"
End Property

Public Property Get QtyReserved() As Long
    On Error GoTo ErrHandler
    QtyReserved = mudtProps.QtyReserved
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyReserved"
End Property
Public Property Get QtyReservedF() As String
    On Error GoTo ErrHandler
    QtyReservedF = mudtProps.QtyReserved
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyReservedF"
End Property
Friend Property Let QtyReserved(val As Long)
    On Error GoTo ErrHandler
    mudtProps.QtyReserved = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyReserved(val)", val
End Property


Public Function SetQtyLastDelivered(val As String) As Boolean
    On Error GoTo ErrHandler
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

    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetQtyLastDelivered(val)", val
End Function
Public Property Get QtyLastOrdered() As Long
    On Error GoTo ErrHandler
    QtyLastOrdered = mudtProps.QtyLastOrdered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastOrdered"
End Property
Public Property Get QtylastSold() As Long
    On Error GoTo ErrHandler
    QtylastSold = mudtProps.QtylastSold
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtylastSold"
End Property
Public Property Get QtyLastOrderedF() As String
    On Error GoTo ErrHandler
    QtyLastOrderedF = mudtProps.QtyLastOrdered
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastOrderedF"
End Property
Friend Property Let QtyLastOrdered(val As Long)
    On Error GoTo ErrHandler
    mudtProps.QtyLastOrdered = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastOrdered(val)", val
End Property
Public Function SetQtyLastOrdered(val As String) As Boolean
    On Error GoTo ErrHandler
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

    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetQtyLastOrdered(val)", val
End Function


Public Property Get RRP() As Long
    On Error GoTo ErrHandler
    RRP = mudtProps.RRP
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.RRP"
End Property
Public Property Get RRPF() As String
    On Error GoTo ErrHandler
    RRPF = Format(mudtProps.RRP / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.RRPF"
End Property
Public Function SetRRP(val As String) As Boolean
    On Error GoTo ErrHandler
Dim lngRRP As Long
    SetRRP = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngRRP = 0
    ElseIf (Not ConvertToLng(val, lngRRP)) Then
        SetRRP = False
        Exit Function
    End If
    mudtProps.RRP = lngRRP
    SetRRP = ValidateObject("RRP")
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetRRP(val)", val
End Function
Public Property Get SP() As Long
    On Error GoTo ErrHandler
    SP = mudtProps.SP
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SP"
End Property
Public Property Get SPF() As String
    On Error GoTo ErrHandler
    SPF = Format(mudtProps.SP / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPF"
End Property
Public Function SetSP(val As String) As Boolean
    On Error GoTo ErrHandler
Dim lngSP As Long
    SetSP = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngSP = 0
    ElseIf (Not ConvertToLng(val, lngSP)) Then
        SetSP = False
        Exit Function
    End If
    mudtProps.SP = lngSP
    SetSP = ValidateObject("SP")
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSP(val)", val
End Function
Public Property Get Cost() As Long
    On Error GoTo ErrHandler
    Cost = mudtProps.Cost
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Cost"
End Property
Public Property Get CostF() As String
    On Error GoTo ErrHandler
    CostF = Format(mudtProps.Cost / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CostF"
End Property
Public Function SetCost(val As String) As Boolean
    On Error GoTo ErrHandler
Dim lngCost As Long
    SetCost = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngCost = 0
    ElseIf (Not ConvertToLng(val, lngCost)) Then
        SetCost = False
        Exit Function
    End If
    mudtProps.Cost = lngCost
    SetCost = ValidateObject("COST")
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCost(val)", val
End Function
Public Property Get SpecialPrice() As Long
    On Error GoTo ErrHandler
    SpecialPrice = mudtProps.SpecialPrice
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialPrice"
End Property
Public Property Get SpecialPriceF() As String
    On Error GoTo ErrHandler
    SpecialPriceF = Format(mudtProps.SpecialPrice / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialPriceF"
End Property
Public Function SetSpecialPrice(val As String) As Boolean
    On Error GoTo ErrHandler
Dim lngCost As Long
    SetSpecialPrice = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngCost = 0
    ElseIf (Not ConvertToLng(val, lngCost)) Then
        SetSpecialPrice = False
        Exit Function
    End If
    mudtProps.SpecialPrice = lngCost
    SetSpecialPrice = ValidateObject("SpecialPrice")
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSpecialPrice(val)", val
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


Public Property Get UKPriceF() As String
    On Error GoTo ErrHandler
    UKPriceF = Format(mudtProps.UKPrice / 100, "#,###.00")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.UKPriceF"
End Property
Public Function SetUKPrice(val As String) As Boolean
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetUKPrice(val)", val
End Function

Public Property Get USPriceF() As String
    On Error GoTo ErrHandler
    USPriceF = Format(mudtProps.USPrice / 100, "$#,###.00")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.USPriceF"
End Property
Public Function SetUSPrice(val As String) As Boolean
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetUSPrice(val)", val
End Function

Public Property Get SeriesTitle() As String
    On Error GoTo ErrHandler
    SeriesTitle = mudtProps.SeriesTitle
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SeriesTitle"
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
    Exit Function
End Function

Public Property Get Author() As String
    On Error GoTo ErrHandler
    Author = Trim(mudtProps.Author)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Author"
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
    Exit Function
End Function

Public Property Get Publisher() As String
    On Error GoTo ErrHandler
    Publisher = Trim(mudtProps.Publisher)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Publisher"
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
    Exit Function
End Function

Public Property Get Availability() As String
    On Error GoTo ErrHandler
    Availability = Trim(mudtProps.Availability)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Availability"
End Property
Public Property Let Availability(p As String)
    On Error GoTo ErrHandler
    mudtProps.Availability = p
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Availability(P)", p
End Property

Public Property Get Status() As String
    On Error GoTo ErrHandler
    Status = FNS(mudtProps.Status)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Status"
End Property
Public Property Get StatusF() As String
    On Error GoTo ErrHandler
    If FNB(mudtProps.Obsolete) = True Then
        StatusF = "OBSOLETE"
    Else
        StatusF = ConvertBookStatus(FNS(mudtProps.Status))
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.StatusF"
End Property
Public Function SetStatus(val As enProductSTatus) As Boolean
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetStatus = True
    Select Case val
    Case enOutOfPrint
        mudtProps.Status = "O"
    Case enAwaitingReprint
        mudtProps.Status = "R"
    Case enInPrint
        mudtProps.Status = "I"
    End Select
    SetStatus = ValidateObject("STATUS")
    mudtProps.IsDirty = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetStatus(val)", val
End Function
Public Property Let ExcludeFromSales(val As Boolean)
    mudtProps.ExcludeFromSales = FNB(val)
End Property
Public Property Get ExcludeFromSales() As Boolean
    ExcludeFromSales = FNB(mudtProps.ExcludeFromSales)
End Property
Public Property Get SupplierID() As Long
    On Error GoTo ErrHandler
    SupplierID = mudtProps.SupplierID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierID"
End Property
Public Property Let SupplierID(val As Long)
    On Error GoTo ErrHandler
    mudtProps.SupplierID = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierID(val)", val
End Property
Public Property Get LastSupplierName() As String
    On Error GoTo ErrHandler
    LastSupplierName = FNS(mudtProps.LastSupplierName)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastSupplierName"
End Property
Public Property Let LastSupplierName(val As String)
    On Error GoTo ErrHandler
    mudtProps.LastSupplierName = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastSupplierName(val)", val
End Property
Public Property Get DEALID() As Long
    On Error GoTo ErrHandler
    DEALID = mudtProps.DEALID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DEALID"
End Property
'Public Property Get DealDiscount() As Double
'    On Error GoTo errHandler
'Dim dblDealDiscount As Double
'Dim oUT As New z_Utility
'    oUT.GetDealDiscount mudtProps.DEALID, dblDealDiscount
'    DealDiscount = dblDealDiscount
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.DealDiscount"
'End Property
Public Property Get LastApproto() As String
    On Error GoTo ErrHandler
    LastApproto = FNS(mudtProps.LastApproto)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastApproto"
End Property
Public Property Get LastDealDescription() As String
    On Error GoTo ErrHandler
    LastDealDescription = FNS(mudtProps.LastDealDescription)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastDealDescription"
End Property
Public Property Get ReturnAvailability() As Integer
    On Error GoTo ErrHandler
    ReturnAvailability = mudtProps.ReturnAvailability
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ReturnAvailability"
End Property

Public Property Get SpecialVat() As Boolean
    On Error GoTo ErrHandler
    SpecialVat = mudtProps.SpecialVat
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialVat"
End Property
Public Property Let SpecialVat(val As Boolean)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SpecialVat = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialVat(val)", val
End Property
Public Property Get IsNONStock() As Boolean
    On Error GoTo ErrHandler
    IsNONStock = (FNS(mudtProps.ProductType) = "N")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsNONStock"
End Property
Public Function SetNONStock()
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "N"
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetNONStock"
End Function
Public Property Get IsGeneralProduct() As Boolean
    On Error GoTo ErrHandler
    IsGeneralProduct = (FNS(mudtProps.ProductType) = "G")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsGeneralProduct"
End Property
Public Function SetGeneralProduct()
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "G"
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetGeneralProduct"
End Function
Public Property Get IsMagsEtc() As Boolean
    On Error GoTo ErrHandler
    IsMagsEtc = (FNS(mudtProps.ProductType) = "M")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsMagsEtc"
End Property
Public Function SetMagsEtc()
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "M"
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetMagsEtc"
End Function
Public Property Get IsBook() As Boolean
    On Error GoTo ErrHandler
    IsBook = (FNS(mudtProps.ProductType) = "B")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsBook"
End Property
Public Function SetBook()
    On Error GoTo ErrHandler
  '  If mcolStack.Count = 0 Then Err.Raise 383
  '  If FNS(mudtProps.ProductType) = "" Then
        mudtProps.ProductType = "B"
        SetDirty True
  '  End If
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetBook"
End Function

Public Property Get ProductType() As String
    On Error GoTo ErrHandler
    ProductType = FNS(mudtProps.ProductType)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ProductType"
End Property
Public Function SetProductType(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = val
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetProductType(val)", val
End Function




Public Property Get Obsolete() As Boolean
    On Error GoTo ErrHandler
    Obsolete = mudtProps.Obsolete
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Obsolete"
End Property
Public Property Let Obsolete(val As Boolean)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Obsolete = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Obsolete(val)", val
End Property
Public Property Get VATRateToUse() As Double
    On Error GoTo ErrHandler
    If SpecialVat Then
        VATRateToUse = VATRate
    Else
        VATRateToUse = oPC.Configuration.VATRate
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRateToUse"
End Property
Public Property Get VATRateToUseF() As String
    On Error GoTo ErrHandler
    If SpecialVat Then
        VATRateToUseF = VATRateF
    Else
        VATRateToUseF = oPC.Configuration.VATRateF
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRateToUseF"
End Property
Public Property Get VATRate() As Double
    On Error GoTo ErrHandler
    VATRate = mudtProps.VATRate
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo ErrHandler
    VATRateF = PBKSPercentF(mudtProps.VATRate)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRateF"
End Property
Public Property Let VATRate(val As Double)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If val <> oPC.Configuration.VATRate Then
        Me.SpecialVat = True
    End If
    mudtProps.VATRate = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRate(val)", val
End Property
Public Function SetVAT(val As String) As Boolean
    On Error GoTo ErrHandler
Dim bOK As Boolean
    
    bOK = SetField_DOUBLE(mudtProps.VATRate, val, "VATRATE")
    If bOK Then
        bOK = ValidateObject("VATRATE")
    End If
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        mudtProps.SpecialVat = True
    End If
    SetVAT = bOK

    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetVAT(val)", val
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
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LastCopySerial = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastCopySerial(val)", val
End Property
Public Property Get LastCopySerial() As Integer
    On Error GoTo ErrHandler
    LastCopySerial = mudtProps.LastCopySerial
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastCopySerial"
End Property
Public Property Let Seesafe(val As Integer)
    On Error GoTo ErrHandler
    mudtProps.PTSeesafe = val
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Seesafe(val)", val
End Property
Public Property Get Seesafe() As Integer
    On Error GoTo ErrHandler
    Seesafe = mudtProps.PTSeesafe
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Seesafe"
End Property

Public Property Let DEALID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsNull(val) Then GoTo EXIT_Handler
    mudtProps.DEALID = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DEALID(val)", val
End Property

'Public Property Get Categories() As z_TextList
'    Set Categories = tlSections
'End Property
'Public Property Get ProductTypes() As z_TextList
'    Set ProductTypes = tlProductTypes
'End Property
'
'Public Property Get CategoryHeadings() As z_TextList
'    Set CategoryHeadings = tlCatalogueHeadings
'End Property

'Public Sub OutOnApproPerTP(ByVal pTPID As Long, ByVal pPID As String, ByRef iQty As Integer, ByRef lngApproID, ByRef lngNumOfApproLines)
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oBatch As z_SQL
'
'
'    Set rs = New ADODB.Recordset
'    GoTo EXIT_Handler
'    MsgBox "Skipping"
'    Set oBatch = New z_SQL
'    lngResult = oBatch.RunGetRecordset("q_ApprosOutstandingPerTitle", enQuery, Array(pPID, pTPID), "", rs)
'    If rs.EOF Then GoTo EXIT_Handler
'
'    Do While Not rs.EOF
'        If lngNumOfApproLines > 1 Then GoTo EXIT_Handler
'        iQty = rs!bal
'        lngApproID = rs!APPL_APP_ID
'        lngNumOfApproLines = lngNumOfApproLines + 1
'        rs.MoveNext
'    Loop
'
'EXIT_Handler:
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
''ERR_Handler:
''    MsgBox Error
''    GoTo EXIT_Handler
''    Resume
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.OutOnApproPerTP(pTPID,pPID,iQty,lngApproID,lngNumOfApproLines)", Array(pTPID, _
'         pPID, iQty, lngApproID, lngNumOfApproLines)
'End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo ErrHandler
Dim oProdCode  As New z_ProdCode
Dim bValid As Boolean
    bValid = True
    If Not (Len(Trim(mudtProps.Title)) > 1 Or Len(Trim(mudtProps.Title)) = 0) Then
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
    oProdCode.LoadNew mudtProps.Code, mudtProps.EAN, mudtProps.ProductType = "B", "", True
    If oProdCode.CodesOK Then
        mobjValid.BreakRule "CODE", False
    Else
        mobjValid.BreakRule "CODE", True
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
    If (FNS(mudtProps.Section) = "") And oPC.Configuration.EnforceSections Then
        mobjValid.BreakRule "SECTION", True
        If pFld = "SECTION" Then bValid = False
    Else
        mobjValid.BreakRule "SECTION", False
    End If

    mobjValid.GetSTatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ValidateObject(pFld)", pFld
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
            strAllRules = mcolClassErrors(Left(pRawErrors, iMarker - 1))
        Else
            strAllRules = mcolClassErrors(pRawErrors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = mcolClassErrors(Mid(pRawErrors, iStart, iMarker - iStart))
            Else
                strRule = mcolClassErrors(Mid(pRawErrors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TranslateErrors(pRawErrors)", pRawErrors
End Function
Private Sub LoadClassErrorsCollection()
    On Error GoTo ErrHandler
    Set mcolClassErrors = New Collection
    mcolClassErrors.Add "Name too short", "NAME"
    mcolClassErrors.Add "Description too short", "TITLE"
    mcolClassErrors.Add "Author too short", "AUTHOR"
    mcolClassErrors.Add "Subtitle too short", "SUBTITLE"
    mcolClassErrors.Add "Publisher too short", "PUBLISHER"
    mcolClassErrors.Add "Publication place too short", "PUBPLACE"
    mcolClassErrors.Add "Publication date too short", "PUBDATE"
    mcolClassErrors.Add "Edition too short", "EDITION"
    mcolClassErrors.Add "Invalid EAN", "EAN"
    mcolClassErrors.Add "Invalid code", "CODE"
    mcolClassErrors.Add "Note too short", "NOTE"
    mcolClassErrors.Add "Invalid cost", "COST"
    mcolClassErrors.Add "Invalid R.R.P.", "RRP"
    mcolClassErrors.Add "Invalid S.P.", "SP"
    mcolClassErrors.Add "Missing section code", "SECTION"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoadClassErrorsCollection"
End Sub
Public Property Get Round() As String
    On Error GoTo ErrHandler
    Round = CStr(mudtProps.PTRound)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Round"
End Property
Public Property Get Discount() As Double
    On Error GoTo ErrHandler
    Discount = mudtProps.PTDiscount
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Discount"
End Property
Public Property Get DiscountF() As String
    On Error GoTo ErrHandler
    DiscountF = PBKSPercentF(mudtProps.PTDiscount)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DiscountF"
End Property

Public Property Get B1Min() As String
    On Error GoTo ErrHandler
    B1Min = CStr(mudtProps.PTB1MIN)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B1Min"
End Property
Public Property Get B1Max() As String
    On Error GoTo ErrHandler
    B1Max = CStr(mudtProps.PTB1Max)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B1Max"
End Property
Public Property Get B1MU() As String
    On Error GoTo ErrHandler
    B1MU = CStr(mudtProps.PTB1MU)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B1MU"
End Property

Public Property Get B2Min() As String
    On Error GoTo ErrHandler
    B2Min = CStr(mudtProps.PTB2MIN)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B2Min"
End Property
Public Property Get B2Max() As String
    On Error GoTo ErrHandler
    B2Max = CStr(mudtProps.PTB2Max)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B2Max"
End Property
Public Property Get B2MU() As String
    On Error GoTo ErrHandler
    B2MU = CStr(mudtProps.PTB2MU)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B2MU"
End Property

Public Property Get B3Min() As String
    On Error GoTo ErrHandler
    B3Min = CStr(mudtProps.PTB3MIN)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B3Min"
End Property
Public Property Get B3Max() As String
    On Error GoTo ErrHandler
    B3Max = CStr(mudtProps.PTB3Max)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B3Max"
End Property
Public Property Get B3MU() As String
    On Error GoTo ErrHandler
    B3MU = CStr(mudtProps.PTB3MU)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B3MU"
End Property

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo ErrHandler
    mudtProps.IsDirty = pVal
 '   RaiseEvent Dirty(pVal)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetDirty(pVal)", pVal
End Sub

Public Property Get Stock6() As Integer
    On Error GoTo ErrHandler
    Stock6 = mudtProps.StckAgeQty6Mnths
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock6"
End Property
Public Property Get Stock12() As Integer
    On Error GoTo ErrHandler
    Stock12 = mudtProps.StckAgeQty12Mnths
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock12"
End Property
Public Property Get Stock18() As Integer
    On Error GoTo ErrHandler
    Stock18 = mudtProps.StckAgeQty18Mnths
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock18"
End Property
Public Property Get Stock18Plus() As Integer
    On Error GoTo ErrHandler
    Stock18Plus = mudtProps.StckAgeQty18MnthsPlus
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock18Plus"
End Property
Public Property Get StockAgedDate() As Integer
    On Error GoTo ErrHandler
    StockAgedDate = mudtProps.StckAgeDate
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.StockAgedDate"
End Property
Public Property Get DateRecordAddedF() As String
    On Error GoTo ErrHandler
    DateRecordAddedF = Format(mudtProps.DateAdded, "dd/mm/yyyy hh:nn")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateRecordAddedF"
End Property
Public Property Get DateLastModifiedF() As String
    On Error GoTo ErrHandler
    If mudtProps.DateLastModified = CDate(0) Then
        DateLastModifiedF = ""
    Else
        DateLastModifiedF = Format(mudtProps.DateLastModified, "dd/mm/yyyy hh:nn")
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastModifiedF"
End Property

Private Function SetField_CURRENCY(fld As Currency, val As String, pValidationName As String)
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_CURRENCY(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_LONG(fld As Long, val As String, pValidationName As String)
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_LONG(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_DATE(fld As Date, val As String, pValidationName As String)
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_DATE(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_STRING(fld As String, val As String, pValidationName As String)
    On Error GoTo ErrHandler
Dim strTemp As String

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_STRING = True
    strTemp = val
    If Len(strTemp) > Len(fld) Then
        Err.Raise vbObjectError + 1001, "String value too long"
    End If
    fld = strTemp
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_STRING(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_DOUBLE(fld As Double, val As String, pValidationName As String)
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_DOUBLE(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function

'Public Sub LoadMovements()
'    On Error GoTo errHandler
'    Set mOSPOs = New c_OSSOrder
'    mOSPOs.Load Me.PID
'    Set mOSCOs = New c_OSCOrder
'    mOSCOs.Load Me.PID
'    Set mOSAPs = New c_OSAPPRO
'    mOSAPs.Load Me.PID
'    Set mMMs = New c_MM
'    mMMs.Load Me.PID
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.LoadMovements"
'End Sub
'Public Sub ReloadMovements(pDate As Date)
'    On Error GoTo errHandler
'    Set mMMs = Nothing
'    Set mMMs = New c_MM
'    mMMs.LoadAll Me.PID, pDate
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.ReloadMovements(pDate)", pDate
'End Sub

'Public Function LoadSales(pYr As Long)
'    On Error GoTo errHandler
'    Set mCurrentSales = New c_SalesPerProduct
'    mCurrentSales.Load Me.PID, pYr, QtyOnHand
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.LoadSales(pYr)", pYr
'End Function
'Public Property Get CurrentSales() As XArrayDB
'    On Error GoTo errHandler
'    Set CurrentSales = mCurrentSales.WeekArray
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.CurrentSales"
'End Property
'Public Property Get CurrentOOS() As XArrayDB
'    On Error GoTo errHandler
'    Set CurrentOOS = mCurrentSales.WeekArrayOOS
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.CurrentOOS"
'End Property
'Public Property Get PreviousSales() As c_SalesPerProduct
'    Set PreviousSales = mPreviousSales
'End Property
'Public Property Get Sections() As z_TextList
'    Set Sections = tlSections
'End Property
