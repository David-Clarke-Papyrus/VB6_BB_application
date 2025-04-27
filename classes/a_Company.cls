VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Company"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As CompanyProps
Private mcolStack As Collection
Private mobjParent As ch_Companies
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Dim pLogo As StdPicture

Event Valid(IsValid As Boolean)
Private mKey As String
Private lngID As Long
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COmpany.IsEditing"
End Property

Public Sub BeginEdit()
    On Error GoTo errHandler
  ' save object state
  mcolStack.Add GetState
  Debug.Print "BeginEdit" & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  Debug.Print "CancelEdit" & mcolStack.Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddCompany Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
  Debug.Print "Applyedit" & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.ApplyEdit"
End Sub

Friend Sub Initialize(Parent As ch_Companies)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set pLogo = New StdPicture
    LoadClassorsCollection
 '   Set cStaff = New c_Staff
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mobjParent Is Nothing Then
  Debug.Print "Term" & mcolStack.Count
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set pLogo = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.IsValid"
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.mobjValid_Status(pMsg)", pMsg
End Sub



Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
Dim udtData As CompanyData
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As CompanyData
  Dim fs As New FileSystemObject
        udtData.buffer = buffer
    LSet mudtProps = udtData
'    If fs.FileExists(oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Logos\" & FNS(mudtProps.LogoFilePath)) Then
'        Set pLogo = LoadPicture(oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Logos\" & FNS(mudtProps.LogoFilePath))
'    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.SetSuperState(buffer)", buffer
End Sub
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.ID"
End Property
Public Property Let ID(val As Long)
    On Error GoTo errHandler
    mudtProps.ID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.ID(val)", val
End Property
Public Property Get CompanyName() As String
    On Error GoTo errHandler
    CompanyName = FNS(mudtProps.CompanyName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CompanyName"
End Property
Public Property Let CompanyName(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CompanyName = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CompanyName(val)", val
End Property
Public Property Get CompanyCode() As String
    On Error GoTo errHandler
    CompanyCode = FNS(mudtProps.CompanyCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CompanyCode"
End Property
Public Property Let CompanyCode(val As String)
   ' On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CompanyCode = val
    If Len(val) > Len(mudtProps.CompanyCode) Then
        Err.Raise 384
    End If
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CompanyCode(val)", val
End Property
Public Property Get VatNumber() As String
    On Error GoTo errHandler
    VatNumber = FNS(mudtProps.VatNumber)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.VatNumber"
End Property
Public Property Let VatNumber(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VatNumber = Trim(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.VatNumber(val)", val
End Property
Public Property Get LogoFilePath() As String
    On Error GoTo errHandler
    LogoFilePath = Trim(mudtProps.LogoFilePath)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.LogoFilePath"
End Property
Public Property Let LogoFilePath(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LogoFilePath = Trim(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.LogoFilePath(val)", val
End Property
Public Property Get StreetandPostalAddress() As String
    On Error GoTo errHandler
    StreetandPostalAddress = RTrim(mudtProps.StreetAddress) & vbCrLf & vbCrLf & RTrim(mudtProps.PostalAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.StreetandPostalAddress"
End Property

Public Property Get StreetAddress() As String
    On Error GoTo errHandler
    StreetAddress = RTrim(mudtProps.StreetAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.StreetAddress"
End Property
Public Property Get StreetAddressDelimited() As String
    On Error GoTo errHandler
    StreetAddressDelimited = PackText(RTrim(mudtProps.StreetAddress))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.StreetAddressDelimited"
End Property
Public Property Let StreetAddress(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StreetAddress = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.StreetAddress(val)", val
End Property
Public Property Get PostalAddress() As String
    On Error GoTo errHandler
    PostalAddress = Trim(mudtProps.PostalAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.PostalAddress"
End Property
Public Property Let PostalAddress(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PostalAddress = Trim(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.PostalAddress(val)", val
End Property
Public Property Get CoRegistrationNumber() As String
    On Error GoTo errHandler
    CoRegistrationNumber = Trim(mudtProps.CoRegistrationNumber)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CoRegistrationNumber"
End Property
Public Property Let CoRegistrationNumber(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CoRegistrationNumber = Trim(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.CoRegistrationNumber(val)", val
End Property
Public Property Get Pastel() As String
    On Error GoTo errHandler
    Pastel = Trim(mudtProps.Pastel)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.Pastel"
End Property
Public Property Let Pastel(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Pastel = Trim(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.Pastel(val)", val
End Property

Public Property Get BankDetails() As String
    On Error GoTo errHandler
    BankDetails = FNS(mudtProps.BankDetails)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.BankDetails"
End Property
Public Property Let BankDetails(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BankDetails = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.BankDetails(val)", val
End Property



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
    ErrorIn "a_Company.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company.LoadClassorsCollection"
End Sub

