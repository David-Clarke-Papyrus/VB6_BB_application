VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_DocumentControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As DCProps
Private mcolStack As Collection
Private mobjParent As ch_DocumentControl
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Event Valid(pMsg As String)
Private mKey As String
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Public Sub BeginEdit()
  mcolStack.Add GetState
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DocumentControl.IsEditing"
End Property

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Public Sub ApplyEdit()
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddDC Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
End Sub

Friend Sub Initialize(Parent As ch_DocumentControl)
  If mudtProps.IsNew Then Set mobjParent = Parent
End Sub

Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    LoadClassorsCollection
End Sub

Private Sub Class_Terminate()
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    
    Set mcolStack = Nothing
    Set mobjValid = Nothing
End Sub

Public Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    RaiseEvent Valid(TranslateErrors(oRS))
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    RaiseEvent Valid(TranslateErrors(oRS))
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid("")
End Sub
Private Sub mobjValid_Status(pMsg As String)
    RaiseEvent Valid(TranslateErrors(pMsg))
End Sub

Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub

Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property
Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty
End Property

Private Function GetState() As String
Dim udtData As DCData
  LSet udtData = mudtProps
  GetState = udtData.buffer
End Function

Private Sub SetState(buffer As String)
Dim udtData As DCData
    udtData.buffer = buffer
    LSet mudtProps = udtData
End Sub
Friend Function GetSuperState() As String
  GetSuperState = GetState
End Function
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get ID() As Long
    ID = mudtProps.ID
End Property
Public Property Let ID(val As Long)
    mudtProps.ID = val
End Property
Public Property Get DOCTypeName() As String
    DOCTypeName = FNS(mudtProps.DOCTypeName)
End Property
Public Property Let DOCTypeName(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCTypeName = FNS(val)
    mudtProps.IsDirty = True
End Property
Public Property Get QtyCopies() As Long
    QtyCopies = mudtProps.QtyCopies
End Property
Public Property Let QtyCopies(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.QtyCopies = val
    mudtProps.IsDirty = True
End Property

Public Property Get Style() As String
    Style = Left(FNS(mudtProps.Style), 1)
End Property
Public Property Get StyleSettings() As String
    StyleSettings = Right(FNS(mudtProps.Style), Len(FNS(mudtProps.Style)) - 1)
End Property
Public Property Get WorkstationID() As Long
    WorkstationID = mudtProps.WSID
End Property
Public Property Let WorkstationID(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.WSID = val
    mudtProps.IsDirty = True
End Property

Public Property Let Style(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Style = val
    mudtProps.IsDirty = True
End Property
Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Get PrinterName()
    PrinterName = FNS(mudtProps.PrinterName)
End Property
Public Function SetPrinter(val As Long, Desc As String)
    mudtProps.PrinterID = val
    mudtProps.PrinterName = Desc
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.IsDirty = True
End Function
Public Function SetPPrintPreview(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    Select Case val
    Case "Print"
        mudtProps.PreviewPrint = "P"
    Case "Preview"
        mudtProps.PreviewPrint = "R"
    End Select
    mudtProps.IsDirty = True
End Function
Public Property Get PreviewPrint() As String
    PreviewPrint = FNS(mudtProps.PreviewPrint)
End Property
Public Property Get PreviewPrintF() As String
    Select Case FNS(mudtProps.PreviewPrint)
    Case "R"
        PreviewPrintF = "Preview"
    Case "P"
        PreviewPrintF = "Print"
    End Select
End Property
Public Property Let PreviewPrint(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PreviewPrint = val
    mudtProps.IsDirty = True
End Property

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
End Function
Private Sub LoadClassorsCollection()
    Set colClassors = New Collection
    colClassors.Add "Missing name", "TYPENAME"
    colClassors.Add "Invalid password", "STYLE"
End Sub





