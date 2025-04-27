VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Deal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As DealProps
Private mcolStack As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mobjParent As ch_Deal
Private colClassors As Collection

Event Valid(pMsg As String)
Dim mKey As String

Private lngID As Long
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
  Debug.Print "BeginEdit" & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.IsEditing"
End Property

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
    ErrorIn "a_Deal.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddDeal Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.ApplyEdit"
End Sub

Friend Sub Initialize(Parent As ch_Deal)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "Description", True
    mobjValid.RuleBroken "Discount", True
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.IsValid"
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
Dim udtData As DealData
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As DealData
        udtData.buffer = buffer
    LSet mudtProps = udtData

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetSuperState(buffer)", buffer
End Sub
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.ID"
End Property
Public Property Let ID(val As Long)
    On Error GoTo errHandler
    mudtProps.ID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.ID(val)", val
End Property
Public Property Get Description() As String
    On Error GoTo errHandler
    Description = FNS(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Description"
End Property
Public Property Let Description(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Description = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Description(val)", val
End Property
Public Function SetDescription(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Description) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Description Then
        mudtProps.Description = val
        SetDirty True
    End If
    SetDescription = True
    Exit Function
End Function
Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = Round(FNDBL(mudtProps.Discount))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Discount"
End Property
Public Property Get DiscountF() As String
    On Error GoTo errHandler
    DiscountF = PBKSPercentF(Round(mudtProps.Discount, 2))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.DiscountF"
End Property
Public Property Let Discount(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Discount = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.Discount(val)", val
End Property
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.Discount
    bOK = SetField_DOUBLE(mudtProps.Discount, val, "DISCOUNT")
    If bOK And (dblTEMP <> mudtProps.Discount) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetDiscount(val)", val
End Function
Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.TPID"
End Property
Public Property Let TPID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.TPID(val)", val
End Property
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "DISCOUNT"
    colClassors.Add "Invalid quantity", "DESCRIPTION"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.LoadClassorsCollection"
End Sub
Private Function TranslateErrors(pRawors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, iStart As Integer
    iMarker = 1
    strAllRules = ""
    pRawors = UCase(pRawors)
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
    ErrorIn "a_Deal.TranslateErrors(pRawors)", pRawors
End Function
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If mudtProps.Description > " " Then
        mobjValid.BreakRule "DESCRIPTION", False
    Else
        mobjValid.BreakRule "DESCRIPTION", True
        If pFld = "DESCRIPTION" Then bValid = False
    End If
    If mudtProps.Discount >= 0 And mudtProps.Discount <= 100 Then
        mobjValid.BreakRule "DISCOUNT", False
    Else
        mobjValid.BreakRule "DISCOUNT", True
        If pFld = "DISCOUNT" Then bValid = False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
'H:
'    MsgBox "CheckOK: " & or
'    Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.ValidateObject(pFld)", pFld
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetDirty(pVal)", pVal
End Sub

Private Function SetField_CURRENCY(fld As Currency, val As String, pValidationName As String)
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_CURRENCY(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_strAsCurrencyToLong(fld As Long, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim cTemp As Currency
Dim bTemp As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_strAsCurrencyToLong = True
    If Trim$(val) = "" Then
        cTemp = 0
    ElseIf Not ConvertToCurr(val, cTemp) Then
        SetField_strAsCurrencyToLong = False
        Exit Function
    End If
    fld = cTemp * oPC.Configuration.DefaultCurrency.Divisor
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_strAsCurrencyToLong(fld,val,pValidationName)", Array(fld, val, _
         pValidationName)
End Function
Private Function SetField_LONG(fld As Long, val As String, pValidationName As String)
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_LONG(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_DATE(fld As Date, val As String, pValidationName As String)
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_DATE(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_STRING(fld As String, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim strTemp As String

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_STRING = True
    strTemp = val
    If Len(strTemp) > Len(fld) Then
        Err.Raise vbObjectError + 1001, "String value too long"
    End If
    fld = strTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_STRING(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_DOUBLE(fld As Double, val As String, pValidationName As String, Optional Factor As Integer)
    On Error GoTo errHandler
Dim dblTEMP As Double
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DOUBLE = True
    If Trim$(val) = "" Then
        dblTEMP = 0
    ElseIf Not ConvertToDBL(val, dblTEMP, Factor) Then
        SetField_DOUBLE = False
        Exit Function
    End If
    fld = dblTEMP
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_DOUBLE(fld,val,pValidationName,Factor)", Array(fld, val, pValidationName, _
         Factor)
End Function
Private Function SetField_DIARYPERIODS(fld As Date, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim dteTemp As Double
Dim bTemp As Boolean
Dim strLeft As String
Dim strRight As String

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DIARYPERIODS = True
    strLeft = UCase(Left(val, 1))
    strRight = UCase(Right(val, 1))
    If strRight <> "W" And strRight <> "M" Then
        SetField_DIARYPERIODS = False
        GoTo EXITH
    End If
    If Not IsNumeric(strLeft) Then
        SetField_DIARYPERIODS = False
        GoTo EXITH
    End If
    If strRight = "W" Then
        fld = DateAdd("ww", CLng(strLeft), Date)
    ElseIf strRight = "M" Then
        fld = DateAdd("m", CLng(strLeft), Date)
    End If
EXITH:
'H:
'    MsgBox error
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.SetField_DIARYPERIODS(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function

Public Sub GetStatus()
    On Error GoTo errHandler
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal.GetSTatus"
End Sub

