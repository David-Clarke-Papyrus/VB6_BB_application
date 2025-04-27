VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Copy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Private mudtProps As CopyProps
Private mcolStack As Collection
Private mobjParent As ch_Copy
Private mOwnerCollection As ch_Copy
Private colClassors As Collection
Private mobjCatalogueEntries As ch_CATALPI
Dim mKey As String
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1


Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As CopyData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As CopyData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
  mudtProps.LocalPrice = mudtProps.Price * oPC.Configuration.LocalCurrency.Factor

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "Catalogues", mobjCatalogueEntries.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String, pOwnerCollection As ch_Copy)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mobjCatalogueEntries.SetSuperState .ReadProperty("Catalogues")
  End With
  Set objPB = Nothing
  Set mOwnerCollection = pOwnerCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetSuperState(buffer,pOwnerCollection)", Array(buffer, pOwnerCollection)
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  mobjCatalogueEntries.ApplyEdit
  If Not mobjParent Is Nothing Then
    mobjParent.AddCopy Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.ApplyEdit"
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
    mobjCatalogueEntries.BeginEdit
     mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
    mobjCatalogueEntries.CancelEdit
    mudtProps.IsDeleted = False
    With mcolStack
        SetState .Item(.Count)
        .Remove .Count
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  Set mobjCatalogueEntries = New ch_CATALPI
  mudtProps.IsNew = True
  mudtProps.PurchaseDate = Date
  LoadClassorsCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_Copy)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
  Set mOwnerCollection = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Initialize(Parent)", Parent
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mobjCatalogueEntries = Nothing
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Class_Terminate"
End Sub

Public Property Get CatalogueEntries() As ch_CATALPI
    Set CatalogueEntries = mobjCatalogueEntries
End Property

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mobjCatalogueEntries.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.IsDirty"
End Property


Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.ID"
End Property
Public Property Get PreviousItem() As a_Copy
    On Error GoTo errHandler
    Set PreviousItem = mOwnerCollection.PreviousItem(Me.ID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PreviousItem"
End Property
Public Property Get CopyBySerialNo(val As Long) As a_Copy
    On Error GoTo errHandler
    Set CopyBySerialNo = mOwnerCollection.FindBySerial(CStr(val))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PreviousItem"
End Property

Public Property Let PID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PID(val)", val
End Property
Public Property Get PID() As Long
    On Error GoTo errHandler
  PID = mudtProps.PID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PID"
End Property
Public Property Get Serial() As Integer
    On Error GoTo errHandler
    Serial = mudtProps.Serial
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Serial"
End Property
Public Property Get SerialF() As String
    On Error GoTo errHandler
    SerialF = "/" & CStr(mudtProps.Serial)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SerialF"
End Property
Public Function SetSerial(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetSerial = True
    mudtProps.Serial = val
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetSerial(val)", val
End Function



Public Property Get Comment() As String
    On Error GoTo errHandler
    Comment = FNS(mudtProps.Comment)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Comment"
End Property
Public Function SetComment(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Comment) Then
        Err.Raise 384
    End If
    SetComment = True
    mudtProps.Comment = val
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get CopyDetails() As String
Dim str As String
    str = "Description: " & Me.Description & vbCrLf & "Comment: " & Me.Comment & vbCrLf & Me.FlagText & vbCrLf & "Price: " & Me.LocalPriceF & vbCrLf & "Purchase date: " & Me.PurchaseDateF & vbCrLf
    CopyDetails = str
End Property
Public Property Get Description() As String
    On Error GoTo errHandler
    Description = Trim$(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Description"
End Property
Public Function SetDescription(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Description) Then
        SetDescription = False
        Err.Raise 384
    End If
    SetDescription = True
    mudtProps.Description = val
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get FlagText() As String
    On Error GoTo errHandler
    FlagText = Trim$(mudtProps.FlagText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.FlagText"
End Property
Public Function SetFlagtext(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.FlagText) Then
        Err.Raise 384
    End If
    SetFlagtext = True
    mudtProps.FlagText = val
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get CatalogueEntries_Concat()
    On Error GoTo errHandler
Dim str As String
Dim i As Integer
    For i = 1 To Me.CatalogueEntries.Count
        str = str & IIf(str > "", ", ", "") & CatalogueEntries(i).Serial & " (" & CatalogueEntries(i).PriceF & ")"
    Next
    CatalogueEntries_Concat = str
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.CatalogueEntries_Concat"
End Property
Public Property Get CatalogueEntriesSimple_Concat()
    On Error GoTo errHandler
Dim str As String
Dim i As Integer
    For i = 1 To Me.CatalogueEntries.Count
        str = str & IIf(str > "", ", ", "") & CatalogueEntries(i).Serial
    Next
    CatalogueEntriesSimple_Concat = str
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.CatalogueEntriesSimple_Concat"
End Property
Public Property Get SoldTo() As String
    On Error GoTo errHandler
    SoldTo = FNS(mudtProps.SoldTo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SoldTo"
End Property


Public Property Get LocalPriceF() As String
    On Error GoTo errHandler
    LocalPriceF = Format(mudtProps.LocalPrice / oPC.Configuration.LocalCurrency.Divisor, oPC.Configuration.LocalCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.LocalPriceF"
End Property
Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Price"
End Property
Public Property Get PriceF() As String
    On Error GoTo errHandler
    If mudtProps.Price > 0 Then
        PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Else
        PriceF = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PriceF"
End Property
Public Property Let Price(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Price = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.Price(val)", val
End Property

Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim lngPrice As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetPrice = True
    If Trim$(val) = "" Then
        lngPrice = 0
    ElseIf Not ConvertToLng(val, lngPrice) Then
        SetPrice = False
        Exit Function
    End If
    If lngPrice < 0 Then
        SetPrice = False
        Exit Function
    End If
 '   If (cPrice <= 10000) And (cPrice Mod 1000 = 0) Then
    If (lngPrice > 0) Then
        mobjValid.RuleBroken "Price", False
    Else
        mobjValid.RuleBroken "Price", True
    End If
    mudtProps.Price = lngPrice
    mudtProps.LocalPrice = mudtProps.Price * oPC.Configuration.LocalCurrency.Factor
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetPrice(val)", val
End Function
Public Property Get PurchaseDate() As Date
    On Error GoTo errHandler
    PurchaseDate = mudtProps.PurchaseDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PurchaseDate"
End Property
Public Property Get PurchaseDateF() As String
    On Error GoTo errHandler
    PurchaseDateF = Format(mudtProps.PurchaseDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PurchaseDateF"
End Property
Public Property Let PurchaseDate(val As Date)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PurchaseDate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.PurchaseDate(val)", val
End Property
Public Function SetPurchaseDate(val As String) As Boolean
    On Error GoTo errHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetPurchaseDate = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetPurchaseDate = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "Purchasedate", False
    Else
        mobjValid.RuleBroken "Purchasedate", True
    End If
    mudtProps.PurchaseDate = dteDate
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetPurchaseDate(val)", val
End Function
Public Property Get SoldDate() As Date
    On Error GoTo errHandler
    SoldDate = mudtProps.SoldDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SoldDate"
End Property
Public Property Get SoldDateF() As String
    On Error GoTo errHandler
    If mudtProps.SoldDate = CDate(0) Then
        SoldDateF = ""
    Else
        SoldDateF = Format(mudtProps.SoldDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SoldDateF"
End Property
Public Property Let SoldDate(val As Date)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SoldDate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SoldDate(val)", val
End Property
Public Function SetSoldDate(val As String) As Boolean
    On Error GoTo errHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetSoldDate = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetSoldDate = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "SoldDate", False
    Else
        mobjValid.RuleBroken "SoldDate", True
    End If
    mudtProps.SoldDate = dteDate
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetSoldDate(val)", val
End Function
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If Not (Len(Trim(mudtProps.Comment)) > 3 Or Len(Trim(mudtProps.Comment)) = 0) Then
        mobjValid.BreakRule "COMMENT", True
        If pFld = "COMMENT" Then bValid = False
    Else
        mobjValid.BreakRule "COMMENT", False
    End If
    If Not (Len(Trim(mudtProps.Comment)) > 3 Or Len(Trim(mudtProps.Comment)) = 0) Then
        mobjValid.BreakRule "Condition", True
        If pFld = "CONDITION" Then bValid = False
    Else
        mobjValid.BreakRule "CONDITION", False
    End If
    If Not (Len(Trim(mudtProps.Price)) > 2) Then
        mobjValid.BreakRule "Price", True
        If pFld = "PRICE" Then bValid = False
    Else
        mobjValid.BreakRule "PRICE", False
    End If

    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.ValidateObject(pVal,pFld)", Array(pVal, pFld)
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
            strAllRules = colClassors(Left(UCase(pRawors), iMarker - 1))
        Else
            strAllRules = colClassors(UCase(pRawors))
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawors, ",")
            If iMarker > 0 Then
                strRule = colClassors(Mid(UCase(pRawors), iStart, iMarker - iStart))
            Else
                strRule = colClassors(Mid(UCase(pRawors), iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Comment too short", "COMMENT"
    colClassors.Add "Description one too short", "DESCRIPTION"
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Purchase date is invalid", "PURCHASEDATE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.LoadClassorsCollection"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.mobjValid_Status(pMsg)", pMsg
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy.SetDirty(pVal)", pVal
End Sub

