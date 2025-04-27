VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COLAllocation"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As COLAllocationProps
Private mcolStack As Collection
Private mobjParent As chex_COLAllocation
Private colClassors As Collection
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
Dim udtData As COLAllocationData

    LSet udtData = mudtProps
    GetState = udtData.buffer
End Function

Private Sub SetState(buffer As String)
Dim udtData As COLAllocationData

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
Dim objPersist As a_COLAllocation_P

    If mcolStack.Count = 0 Then Err.Raise 445
    
    If Not mobjParent Is Nothing Then
        mobjParent.AddCOLAllocation Me
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

Public Property Get CreateInvoice() As Boolean
    CreateInvoice = mudtProps.CreateInvoice
End Property

Friend Sub Initialize(Parent As chex_COLAllocation)
  Set mobjParent = Parent
End Sub

Private Sub Class_Initialize()

    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set colClassors = New Collection
    LoadClassorsCollection
End Sub

Private Sub Class_Terminate()
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
    
  Set mcolStack = Nothing
  Set colClassors = Nothing
End Sub

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

Public Property Get ID() As Long
    ID = mudtProps.ID
End Property

Public Property Get CustomerBlocked() As Boolean
    CustomerBlocked = FNB(mudtProps.CustomerBlocked)
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
Public Property Get DeliveredSoFarSS() As Long
    DeliveredSoFarSS = mudtProps.DeliveredSoFarSS
End Property
Public Property Get DeliveredSoFarFirm() As Long
    DeliveredSoFarFirm = mudtProps.DeliveredSoFarFirm
End Property
Public Property Get ProductOH() As Long
    ProductOH = mudtProps.ProductOH
End Property
Public Property Get ProductRES() As Long
    ProductRES = mudtProps.ProductRES
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
Public Property Get QtyonCO() As Long
    QtyonCO = mudtProps.QtyonCO
End Property
Public Property Get QtyonOrder() As Long
    QtyonOrder = mudtProps.QtyonOrder
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
Public Property Get OrderedSSQTY() As Long
    OrderedSSQTY = mudtProps.OrderedSSQTY
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
Public Property Get CustomerAcnoName() As String
    If oPC.GetProperty("ShowCustomerAcnoLeftOfName") = "TRUE" Then
        CustomerAcnoName = IIf(CustomerAcno > "", "(", "") & CustomerAcno & IIf(CustomerAcno > "", ")", "") & CustomerName
    Else
        CustomerAcnoName = CustomerName & IIf(CustomerAcno > "", "(", "") & CustomerAcno & IIf(CustomerAcno > "", ")", "")
    End If
End Property
Public Property Get CustomerInitials() As String
    CustomerInitials = FNS(mudtProps.CustomerInitials)
End Property
Public Property Get CustomerTitle() As String
    CustomerTitle = FNS(mudtProps.CustomerTitle)
End Property
Public Property Get CustomerAcno() As String
    CustomerAcno = FNS(mudtProps.CustomerAcno)
End Property
Public Property Get DepositValueF() As String
    DepositValueF = Format(mudtProps.DepositValue, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get DepositValue() As Double
    DepositValue = mudtProps.DepositValue
End Property

Public Property Get Fullname() As String
    Fullname = CustomerName & IIf(Len(Title) > 0 Or Len(CustomerInitials) > 0, ", ", "") & CustomerTitle & IIf(Len(CustomerTitle) > 0, " ", "") & CustomerInitials
End Property
Public Property Get Note() As String
    Note = FNS(mudtProps.Note)
End Property

Public Property Get Phone() As String
    Phone = FNS(mudtProps.Phone)
End Property
Public Property Get ItemCode() As String
    ItemCode = FNS(mudtProps.code)
End Property
Public Property Get UsesSubstitutions() As String
    UsesSubstitutions = FNS(mudtProps.UsesSubstitutesYN)
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
Public Property Get CodeTitleShort(pLen As Integer) As String
Dim strTmp As String

    strTmp = FNS(mudtProps.Title) & " " & FNS(mudtProps.code)
    If Len(strTmp) - pLen < 3 And Len(strTmp) - pLen > 0 Then
        strTmp = Left(strTmp, Len(strTmp) - 3) & "..."
    ElseIf Len(strTmp) - pLen > 3 Then
        strTmp = Left(strTmp, pLen - 3) & "..."
    End If
    CodeTitleShort = strTmp
    
End Property
Public Property Get code() As String
Dim strTmp As String


    code = FNS(mudtProps.code)

End Property
Public Property Get OrderDetails() As String
    OrderDetails = OrderCode & "(" & OrderDateF & "):" & Ref
End Property
Public Property Get WSLock() As String
    WSLock = FNS(mudtProps.WSLock)
End Property
Public Property Get AllocatedQty() As Long
    AllocatedQty = mudtProps.AllocatedQty
End Property
Public Property Let AllocatedQty(val As Long)
    If val <> mudtProps.AllocatedQty Then
        mudtProps.AllocatedQty = val
        SetDirty True
    End If
End Property
Public Function SetAllocatedQty(pQty As Long) As Boolean
    mudtProps.AllocatedQty = pQty
End Function
Public Property Get AllocatedQtySS() As Long
    AllocatedQtySS = mudtProps.AllocatedQtySS
End Property
Public Property Let AllocatedQtySS(val As Long)
    If val <> mudtProps.AllocatedQtySS Then
        mudtProps.AllocatedQtySS = val
        SetDirty True
    End If
End Property
Public Function SetAllocatedQtySS(pQty As Long) As Boolean
    mudtProps.AllocatedQtySS = pQty
End Function

Public Function SetACtionYN(pOn As Boolean)
    mudtProps.ActionYN = pOn
End Function
Public Property Get ActionYN() As Boolean
    ActionYN = mudtProps.ActionYN
End Property
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
    colClassors.Add "Invalid allocation quantity", "ALLOCATEDQTY"
End Sub
Private Sub SetDirty(pVal As Boolean)
    mudtProps.IsDirty = pVal
   ' RaiseEvent Dirty(pVal)
End Sub

