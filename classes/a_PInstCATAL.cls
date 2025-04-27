VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PInstCATAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Event Valid(Msg As String)
Private mudtProps As PICATALProps
Private mcolStack As Collection
Private mobjParent As ch_PICATAL
Private colClassErrors As Collection

Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private Function GetState() As String
  Dim mudtData As PICATALData
  
  LSet mudtData = mudtProps
  GetState = mudtData.Buffer
End Function

Private Sub SetState(Buffer As String)
  Dim mudtData As PICATALData
  
  mudtData.Buffer = Buffer
  LSet mudtProps = mudtData
End Sub
Friend Function GetSuperState() As String
  GetSuperState = GetState
End Function
Friend Sub SetSuperState(Buffer As String)
  SetState Buffer
End Sub

Public Sub ApplyEdit()
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddPICATAL Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
End Sub

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

Private Sub Class_Initialize()
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
End Sub
Friend Sub Initialize(Parent As ch_PICATAL)
  If mudtProps.IsNew Then Set mobjParent = Parent
End Sub

Friend Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property

Private Sub Class_Terminate()
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mcolStack = Nothing
End Sub

Private Sub objValid_BrokenRule(errors As String)
    RaiseEvent Valid(errors)
End Sub

Private Sub objValid_NoBrokenRules()
    RaiseEvent Valid("")
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


Public Property Get ID() As Long
    ID = mudtProps.ID
End Property

Public Property Let PIID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PIID = val
    mudtProps.IsDirty = True
End Property
Public Property Get PIID() As Long
  PIID = mudtProps.PIID
End Property
Public Property Get CATALID() As String
    CATALID = FixNullsString(mudtProps.CATALID)
End Property
Public Property Let CATALID(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CATALID = val
    mudtProps.IsDirty = True
End Property

Public Property Get Serial() As Integer
    Serial = mudtProps.Serial
End Property
Public Property Let Serial(val As Integer)
    mudtProps.Serial = val
End Property


Public Property Get Price() As Long
    Price = mudtProps.Price
End Property
Public Property Get PriceF() As String
    PriceF = Format(mudtProps.Price / 100, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Let Price(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Price = val
    mudtProps.IsDirty = True
End Property
Public Function SetPrice(val As String) As Boolean
Dim lngPrice As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetPrice = True
    If Trim$(val) = "" Then
        lngPrice = 0
    ElseIf Not ConvertToLng(val, lngPrice) Then
        SetPrice = False
        Exit Function
    End If
    If lngPrice > 0 Then
        mobjValid.RuleBroken "Price", False
    Else
        mobjValid.RuleBroken "Price", True
    End If
    mudtProps.Price = lngPrice
    mudtProps.IsDirty = True
End Function
Private Function CheckInput(pVal As String, pFld As String) As Boolean

    CheckInput = True
    If mudtProps.Price < 0 Then
        mobjValid.BreakRule "PRICE", True
        If pFld = "PRICE" Then CheckInput = False
    Else
        mobjValid.BreakRule "PRICE", False
    End If
    mobjValid.GetStatus
End Function
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
Private Sub LoadClassErrorsCollection()
    Set colClassErrors = New Collection
    colClassErrors.Add "Invalid price", "PRICE"
End Sub



