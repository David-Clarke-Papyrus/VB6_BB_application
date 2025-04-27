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
Attribute VB_Exposed = True
Option Explicit
Private mudtProps As ConfigProps
Private mcolStack As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Event Valid(Errors As String, Status As Boolean)

Private colClassErrors As Collection
Private strErrors As String

Private Function GetState() As String
  Dim udtData As ConfigData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
End Function

Private Sub SetState(Buffer As String)
  Dim udtData As ConfigData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData
End Sub

Public Function GetSuperState() As String
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Public Sub SetSuperState(Buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = Buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
  End With
  Set objPB = Nothing
End Sub

Public Function Load() As Long

    Dim objPersist As a_Configuration_P
    
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_Configuration_P
    SetSuperState objPersist.Fetch()
    Set objPersist = Nothing
End Function
Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set colClassErrors = New Collection
    LoadClassErrorsCollection
    mudtProps.IsNew = True
End Sub

Private Sub Class_Terminate()
    If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set colClassErrors = Nothing
End Sub
Private Sub LoadClassErrorsCollection()
    Set colClassErrors = New Collection
    colClassErrors.Add "Test", "ERRTEST"
End Sub
Private Function TranslateErrors(pRawErrors As String) As String
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
End Function

Private Sub mobjValid_BrokenRule(pErrors As String)
    strErrors = TranslateErrors(pErrors)
    RaiseEvent Valid(TranslateErrors(pErrors), IsValid)
End Sub
Private Sub mobjValid_RuleUnbroken(pErrors As String)
    strErrors = TranslateErrors(pErrors)
    RaiseEvent Valid(TranslateErrors(pErrors), IsValid)
End Sub

Private Sub mobjValid_NoBrokenRules()
    strErrors = ""
    RaiseEvent Valid("", True)
End Sub

Public Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property
Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = (mcolStack.Count > 0)
End Property

Public Sub BeginEdit()
  mcolStack.Add GetState
End Sub
Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub
Public Sub ApplyEdit(pErrors As String)
Dim objPersist As a_Configuration_P
Dim strStatus As String

    If mcolStack.Count = 0 Then Err.Raise 445
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
    pErrors = ""
    mudtProps.IsDirty = False
End Sub

Public Property Get AS400COnnectionString() As String
    AS400COnnectionString = FNS(mudtProps.AS400COnnectionString)
End Property
Public Property Get Q() As String
    Q = FNS(mudtProps.Q)
End Property
Public Property Get QQ() As String
    QQ = FNS(mudtProps.QQ)
End Property
Public Property Get BelowDIP() As String
    BelowDIP = FNS(mudtProps.BelowDIP)
End Property
Public Property Get GrowerFolder() As String
    GrowerFolder = FNS(mudtProps.GrowerFolder)
End Property
Public Property Get RepFolder() As String
    RepFolder = FNS(mudtProps.RepFolder)
End Property
Public Property Get PDFFolder() As String
    PDFFolder = FNS(mudtProps.PDFFolder)
End Property
Public Property Get FTPAddress() As String
    FTPAddress = FNS(mudtProps.FTPAddress)
End Property
Public Property Get FTPUsername() As String
    FTPUsername = FNS(mudtProps.FTPUsername)
End Property
Public Property Get FTPPassword() As String
    FTPPassword = FNS(mudtProps.FTPPassword)
End Property
Public Property Get FTPDefaultFolder() As String
    FTPDefaultFolder = FNS(mudtProps.FTPDefaultFolder)
End Property
Public Property Get FTPPassive() As Boolean
    FTPPassive = FNB(mudtProps.FTPPassive)
End Property

Public Property Let AS400COnnectionString(Val As String)
    mudtProps.AS400COnnectionString = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let Q(Val As String)
    mudtProps.Q = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let QQ(Val As String)
    mudtProps.QQ = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let BelowDIP(Val As String)
    mudtProps.BelowDIP = FNS(Val)
    mudtProps.IsDirty = True
End Property

Public Property Let GrowerFolder(Val As String)
    mudtProps.GrowerFolder = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let RepFolder(Val As String)
    mudtProps.RepFolder = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let PDFFolder(Val As String)
    mudtProps.PDFFolder = FNS(Val)
    mudtProps.IsDirty = True
End Property

Public Property Let FTPUsername(Val As String)
    mudtProps.FTPUsername = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let FTPPassword(Val As String)
    mudtProps.FTPPassword = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let FTPAddress(Val As String)
    mudtProps.FTPAddress = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let FTPDefaultFolder(Val As String)
    mudtProps.FTPDefaultFolder = FNS(Val)
    mudtProps.IsDirty = True
End Property
Public Property Let FTPPassive(Val As Boolean)
    mudtProps.FTPPassive = FNB(Val)
    mudtProps.IsDirty = True
End Property


