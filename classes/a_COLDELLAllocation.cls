VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COLDELLAllocation"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As COLDELLAllocationProps
Private mcolStack As Collection
Private mobjParent As chex_COLDELLAllocation
Private colClassErrors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Event Valid(strMsg As String)
Dim mKey As String

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
'Public Property Let Manager(pManager As chex_Commission)
'    Set oManager = pManager
'End Property
Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub

Private Function GetState() As String
Dim udtData As COLDELLAllocationData

    LSet udtData = mudtProps
    GetState = udtData.buffer
End Function

Private Sub SetState(buffer As String)
Dim udtData As COLDELLAllocationData

    udtData.buffer = buffer
    LSet mudtProps = udtData
End Sub

Public Function GetSuperState() As String
  GetSuperState = GetState
End Function

Public Sub SetSuperState(buffer As String)
  SetState buffer
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

Public Sub ApplyEdit()
Dim objPersist As a_COLDELLAllocation_P

    If mcolStack.Count = 0 Then Err.Raise 445
    
    If Not mobjParent Is Nothing Then
        mobjParent.AddCOLDELLAllocation Me
        Set mobjParent = Nothing
    End If
    
    mcolStack.Remove mcolStack.Count
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

Public Property Get IsValid() As Boolean
    IsValid = (mobjValid.Count = 0)
End Property
Friend Sub Initialize(Parent As chex_COLDELLAllocation)
  Set mobjParent = Parent
End Sub

Private Sub Class_Initialize()

    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set colClassErrors = New Collection
    LoadClassErrorsCollection
End Sub

Private Sub Class_Terminate()
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
    
  Set mcolStack = Nothing
  Set colClassErrors = Nothing
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    RaiseEvent Valid(TranslateErrors(errors))
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    RaiseEvent Valid(TranslateErrors(errors))
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid("")
End Sub
Private Sub mobjValid_Status(pMsg As String)
    RaiseEvent Valid(TranslateErrors(pMsg))
End Sub

Public Property Get ID() As Long
    ID = mudtProps.ID
End Property


Public Property Get COLID() As Long
    COLID = mudtProps.COLID
End Property
Public Property Let COLID(val As Long)
    mudtProps.COLID = val
    mudtProps.IsDirty = True
End Property
Public Property Get DELLID() As Long
    DELLID = mudtProps.DELLID
End Property
Public Property Let DELLID(val As Long)
    mudtProps.DELLID = val
    mudtProps.IsDirty = True
End Property
Public Property Get PID() As String
    PID = FNS(mudtProps.PID)
End Property
Public Property Let PID(val As String)
    mudtProps.PID = val
    mudtProps.IsDirty = True
End Property
Public Property Get DeliveredSoFar() As Long
    DeliveredSoFar = mudtProps.DeliveredSoFar
End Property
Public Property Get QtyOnHand() As Long
    QtyOnHand = mudtProps.QtyOnHand
End Property
Public Property Get QtyOS() As Long
    QtyOS = mudtProps.OrderedQty - mudtProps.DeliveredSoFar
End Property
Public Property Get QtyReserved() As Long
    QtyReserved = mudtProps.QtyReserved
End Property

Public Property Get QtyJustReceived() As Long
    QtyJustReceived = mudtProps.QtyJustReceived
End Property
Public Property Get QuotedPrice() As Long
    QuotedPrice = mudtProps.QuotedPrice
End Property
Public Property Get OrderedQty() As Long
    OrderedQty = mudtProps.OrderedQty
End Property
Public Property Get Status() As Long
    Status = mudtProps.Status
End Property
Public Property Get OrderDate() As Date
    OrderDate = mudtProps.OrderDate
End Property
Public Property Get OrderDateF() As String
    OrderDateF = Format(mudtProps.OrderDate, "dd/mm/yyyy")
End Property
Public Property Get Ref() As String
    Ref = FNS(mudtProps.Ref)
End Property
Public Property Get CustomerName() As String
    CustomerName = FNS(mudtProps.CustomerName)
End Property
Public Property Get OrderCode() As String
    OrderCode = FNS(mudtProps.OrderCode)
End Property
Public Property Get Title() As String
    Title = FNS(mudtProps.Title)
End Property
Public Property Get TitleShort(pLen As Integer) As String
    TitleShort = Left(FNS(mudtProps.Title), pLen) & IIf(Len(FNS(mudtProps.Title)) > pLen, "...", "")
End Property

Public Property Get OrderDetails() As String
    OrderDetails = OrderCode & "(" & OrderDateF & "):" & Ref
End Property

Public Property Get AllocatedQty() As Long
    AllocatedQty = mudtProps.AllocatedQty
End Property
Public Property Let AllocatedQty(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    
  '  mobjValid.RuleBroken "AllocatedQty", val <= 0 Or val > 1000
    If val <> mudtProps.AllocatedQty Then
        mudtProps.AllocatedQty = val
        SetDirty True
    End If
End Property
Public Function SetAllocatedQty(pQty As Long) As Boolean
    AllocatedQty = pQty
  '  mobjParent.GetSTatus
End Function

Public Property Get IsEditing() As Boolean
    IsEditing = mcolStack.Count > 0
End Property
Private Function CheckInput(pFld As String) As Boolean

    CheckInput = True
    If (mudtProps.AllocatedQty < 1 Or mudtProps.AllocatedQty > 1000) Then
        mobjValid.BreakRule "AllocatedQty", True
        If pFld = "AllocatedQty" Then CheckInput = False
    Else
        mobjValid.BreakRule "AllocatedQty", False
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
    colClassErrors.Add "Invalid allocation quantity", "ALLOCATEDQTY"
End Sub
Private Sub SetDirty(pVal As Boolean)
    mudtProps.IsDirty = pVal
   ' RaiseEvent Dirty(pVal)
End Sub

