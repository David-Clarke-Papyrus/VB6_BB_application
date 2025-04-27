VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------------------------------------

Event Valid(msg As String)
Private mudtProps As PaymentProps
Private mcolStack As Collection
Private mobjParent As ch_Payment
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassErrors As Collection
Private mExchangeparent As a_Exchange

Private mKey As String

'Private mLPLessDisc As Long
'
'Private mLPLessDiscExt As Long
'
'Private mLPDiscExt As Long
'
'Private mLPLessDiscExtNonVAT As Long
'
'Private mLPLessDiscExtVAT As Long
'
'Private mLPDisc_Unit As Long
'
'Private mLPDisc_Ext As Long
'
Public Function CalculateLine()

'    mLPLessDisc = mudtProps.Price * ((100 - mudtProps.DiscountRate) / 100)
'    mLPLessDiscExt = mLPLessDisc * mudtProps.Qty
'    mLPDiscExt = (mudtProps.Price * mudtProps.Qty) - mLPLessDiscExt
'    mLPLessDiscExtVAT = (mLPLessDiscExt - (mLPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
'    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT
'
End Function

Private Function GetState() As String
  Dim mudtData As PaymentData
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
End Function
    
Private Sub SetState(buffer As String)
  Dim mudtData As PaymentData
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Function GetSuperState() As String
    GetSuperState = GetState
End Function
Friend Sub SetSuperState(buffer As String)
    SetState buffer
End Sub

Public Sub ApplyEdit()
    If mcolStack.Count = 0 Then Err.Raise 445
    If Not mobjParent Is Nothing Then
        mobjParent.AddPaymentLine Me
        Set mobjParent = Nothing
    End If
    mcolStack.Remove mcolStack.Count
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
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

Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassErrorsCollection
    mobjValid.RuleBroken "AMT", True
    mudtProps.IsNew = True
End Sub
Friend Sub Initialize(Parent As ch_Payment)
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
End Sub
Friend Sub SetParentExchange(pParentExchange As a_Exchange)
    Set mExchangeparent = pParentExchange
End Sub
Friend Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property

Private Sub Class_Terminate()
 '   Set moProd = Nothing
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    Set mcolStack = Nothing
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

Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub
Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = (mcolStack.Count > 0)
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty
End Property


Public Property Get PaymentID() As Long
    PaymentID = mudtProps.PaymentID
End Property

Public Property Let PaymentID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PaymentID = val
    mudtProps.IsDirty = True
End Property
Public Property Get ExchangeID() As String
  ExchangeID = mudtProps.EXCHANGE_GUID
End Property


Public Function SetType(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Type) Then
        Err.Raise 384
    End If
    If val <> FNS(mudtProps.Type) Then
        mudtProps.Type = val
        SetDirty True
    End If
    SetType = True
End Function
Public Function SetReference(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Reference) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Reference Then
        mudtProps.Reference = val
        SetDirty True
    End If
    SetReference = True
End Function
Public Property Get ReferenceComplete() As String
    ReferenceComplete = FNS(mudtProps.Reference)
End Property
Public Property Get ReferenceCode() As String
    ReferenceCode = Left(FNS(mudtProps.Reference), 1)
End Property
Public Property Get ReferenceDescription() As String
Dim str As String
    str = FNS(mudtProps.Reference)
    ReferenceDescription = Right(str, Len(str) - 2)
End Property
Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Note Then
        mudtProps.Note = val
        SetDirty True
    End If
    SetNote = True
End Function
Public Property Get Note() As String
    Note = FNS(mudtProps.Note)
End Property

Public Property Get VoucherType() As String
    VoucherType = Left(FNS(mudtProps.Reference), 1)
End Property

Public Property Get Amt() As Long
    Amt = FNN(mudtProps.Amt)
End Property
Public Property Get AmtF() As String
    AmtF = Format(mudtProps.Amt / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property
Public Property Get AmtF_nonNegative() As String
Dim tmp As Long
    tmp = mudtProps.Amt
    If tmp < 0 Then
        tmp = tmp * -1
    End If
    AmtF_nonNegative = Format(tmp / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property

Public Property Let Amt(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Amt Then
        mudtProps.Amt = val
        SetDirty True
    End If
    ValidateObject "Amt"
End Property
Public Function SetAmt(val As String) As Boolean
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Amt
    bOK = SetField_strAsCurrencyToLong(mudtProps.Amt, val, mcolStack.Count, "Amt", False, oPC.CurrencyDivisor)
    If bOK And (lngTemp <> mudtProps.Amt) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("Amt")
    End If
    
    SetAmt = bOK
End Function
Public Property Let COLID(val As Long)
    mudtProps.COLID = val
End Property
Public Property Get COLID() As Long
    COLID = FNN(mudtProps.COLID)
End Property

Public Property Get PaymentType() As String
    PaymentType = Trim(mudtProps.Type)
End Property
Public Property Get PaymentTypeF() As String
    Select Case UCase(Trim(mudtProps.Type))
    Case "C"
        PaymentTypeF = "Cash"
    Case "A"
        PaymentTypeF = "Credit card"
    Case "V"
        PaymentTypeF = "V. redeemed"
    Case "Q"
        PaymentTypeF = "Cheque"
    Case "RD"
        PaymentTypeF = "Redeemed deposit"
    Case "DDP"
        PaymentTypeF = "Direct deposit"
    Case "CN"
        PaymentTypeF = "Credit note"
    Case "CV"
        PaymentTypeF = "Credit voucher"
    Case "CNR"
        PaymentTypeF = "Credit voucher as change"
    Case "AC"
        PaymentTypeF = "On account"
    End Select
End Property

Private Function TranslateErrors(pRawErrors As String) As String
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, istart As Integer
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
            istart = iMarker + 1
            iMarker = InStr(istart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = colClassErrors(MID(pRawErrors, istart, iMarker - istart))
            Else
                strRule = colClassErrors(MID(pRawErrors, istart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
End Function

Private Sub SetDirty(pVal As Boolean)
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
End Sub

Private Sub LoadClassErrorsCollection()
    Set colClassErrors = New Collection
    colClassErrors.Add "Invalid price", "AMT"
    colClassErrors.Add "Invalid payment type", "TYPE"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
On Error GoTo ERRH
Dim bValid As Boolean
    bValid = True
    If mudtProps.Amt > 0 Then
        mobjValid.RuleBroken "AMT", False
    Else
        mobjValid.RuleBroken "AMT", True
        If pFld = "AMT" Then bValid = False
    End If
    If mudtProps.Type = "Q" Or mudtProps.Type = "C" Or mudtProps.Type = "CC" Or mudtProps.Type = "V" Then
        mobjValid.RuleBroken "TYPE", False
    Else
        mobjValid.RuleBroken "TYPE", True
        If pFld = "TYPE" Then bValid = False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
    Exit Function
ERRH:
    MsgBox "CheckOK: " & Error
End Function

Public Sub GetStatus()
    ValidateObject ""
    mobjValid.GetStatus
End Sub

