VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Catalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(strMsg As String)
Event dbactionStatus(Status As Integer)

Private mudtProps As CatalogueProps
Private mcolStack As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As CatalogueData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
Dim udtData As CatalogueData
    udtData.buffer = buffer
    LSet mudtProps = udtData
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.SetState(buffer)", buffer
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
    ErrorIn "a_Catalogue.GetSuperState"
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
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.SetSuperState(buffer)", buffer
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.BeginEdit"
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
    ErrorIn "a_Catalogue.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim objPersist As a_Catalogue_P
  
  If mcolStack.Count = 0 Then Err.Raise 445

  Set objPersist = New a_Catalogue_P
  If mudtProps.IsDeleted Then
    ' code to delete the object's data goes here
    objPersist.DeleteObject mudtProps.ID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    ' save object to database if appropriate
    ' save object state
    mcolStack.Remove mcolStack.Count
    SetSuperState objPersist.Save(GetSuperState)
    If mudtProps.dbactionStatus = 22 Then RaiseEvent dbactionStatus(22)
    mudtProps.IsNew = False
  Else
    mcolStack.Remove mcolStack.Count
  End If
  Set objPersist = Nothing
  mudtProps.IsDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.ApplyEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set colClassors = New Collection
    LoadClassorsCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Class_Terminate"
End Sub
Public Sub Delete()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    mudtProps.IsDeleted = True
    mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Delete"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.mobjValid_Status(pMsg)", pMsg
End Sub
Public Function Load(CatalogueID As Long) As Boolean
    On Error GoTo errHandler
Dim objPersist As a_Catalogue_P
  
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    ' code to load the object goes here
    Set objPersist = New a_Catalogue_P
    Load = True
    SetSuperState objPersist.Fetch(CatalogueID)
    Set objPersist = Nothing
    mobjValid.RuleBroken "Description", False
    mobjValid.RuleBroken "Serial", False
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Load(CatalogueID)", CatalogueID
End Function

Public Property Get CatalogueID() As Long
    On Error GoTo errHandler
  CatalogueID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.CatalogueID"
End Property


Public Property Get Description() As String
    On Error GoTo errHandler
  Description = Trim$(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Description"
End Property
Public Function SetDescription(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetDescription = True
    If Len(val) > Len(mudtProps.Description) Then
        Err.Raise 384
    End If
    mudtProps.Description = val
    SetDescription = ValidateObject(val, "DESCRIPTION")
    mudtProps.IsDirty = True
    Exit Function
End Function

'((((((((((((((((
Public Property Get Serial() As Integer
    On Error GoTo errHandler
    Serial = mudtProps.Serial
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Serial"
End Property
Public Property Let Serial(val As Integer)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Serial = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Serial(val)", val
End Property
Public Function SetSerial(val As String) As Boolean
    On Error GoTo errHandler
Dim iSerial As Integer
  If mcolStack.Count = 0 Then Err.Raise 383
    SetSerial = True
    If Trim$(val) = "" Then
        iSerial = 0
    ElseIf Not ConvertToInt(val, iSerial) Then
        SetSerial = False
        Exit Function
    End If
    If (iSerial <= 10000) And (iSerial Mod 1000 = 0) Then
        mobjValid.RuleBroken "SERIAL", False
    Else
        mobjValid.RuleBroken "SERIAL", True
    End If
    mudtProps.Serial = iSerial
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.SetSerial(val)", val
End Function


Public Property Get DateFirstPrinted() As Date
    On Error GoTo errHandler
    DateFirstPrinted = mudtProps.DateFirstPrinted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.DateFirstPrinted"
End Property
Public Property Get DateFirstPrintedF() As String
    On Error GoTo errHandler
    DateFirstPrintedF = Format(mudtProps.DateFirstPrinted, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.DateFirstPrintedF"
End Property
Public Property Let DateFirstPrinted(val As Date)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DateFirstPrinted = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.DateFirstPrinted(val)", val
End Property
Public Function SetDateFirstPrinted(val As String) As Boolean
    On Error GoTo errHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDateFirstPrinted = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDateFirstPrinted = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateFirstPrinted", False
    Else
        mobjValid.RuleBroken "DateFirstPrinted", True
    End If
    mudtProps.DateFirstPrinted = dteDate
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.SetDateFirstPrinted(val)", val
End Function

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.IsNew"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.IsEditing"
End Property
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.IsValid"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.IsDirty"
End Property
Public Function Validate(pMsg As String) As Boolean
    On Error GoTo errHandler
    pMsg = mobjValid.AllBrokenRules
    Validate = (pMsg = "")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.Validate(pMsg)", pMsg
End Function
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If Not mudtProps.Serial > 0 Then
        mobjValid.BreakRule "serial", True
        If pFld = "SERIAL" Then bValid = False
    Else
        mobjValid.BreakRule "SERIAL", False
    End If
    If Not (Len(Trim(mudtProps.Description)) > 2) Then
        mobjValid.BreakRule "DESCRIPTION", True
        If pFld = "DESCRIPTION" Then bValid = False
    Else
        mobjValid.BreakRule "DESCRIPTION", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.ValidateObject(pVal,pFld)", Array(pVal, pFld)
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
    ErrorIn "a_Catalogue.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Description too short", "DESCRIPTION"
    colClassors.Add "Serial number missing", "SERIAL"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.LoadClassorsCollection"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue.SetDirty(pVal)", pVal
End Sub

