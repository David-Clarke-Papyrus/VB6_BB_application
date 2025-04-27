VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(strMsg As String)
Event dbactionStatus(Status As Integer)
Event InUse()
Private mudtProps As PTProps
Private mcolStack As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As PTData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
Dim udtData As PTData
    udtData.buffer = buffer
    LSet mudtProps = udtData
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetState(buffer)", buffer
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
    ErrorIn "a_PT.GetSuperState"
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
    ErrorIn "a_PT.SetSuperState(buffer)", buffer
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.BeginEdit"
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
    ErrorIn "a_PT.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim objPersist As a_PT_P
  
  If mcolStack.Count = 0 Then Err.Raise 445

  Set objPersist = New a_PT_P
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
    ErrPreserve
            
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.ApplyEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    mudtProps.B1Min = 0
    mudtProps.B1Max = 10000
    mudtProps.B2Min = 10000
    mudtProps.B2Max = 50000
    mudtProps.B3Min = 50000
    mudtProps.B3Max = 1000000
    mudtProps.Active = True
    Set colClassors = New Collection
    LoadClassorsCollection
    ValidateObject ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Class_Terminate"
End Sub
Public Sub Delete()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    mudtProps.IsDeleted = True
    mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Delete"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
Dim s As String

    On Error GoTo errHandler
    s = TranslateErrors(pMsg)
    RaiseEvent Valid(s)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.mobjValid_Status(pMsg)", pMsg
End Sub
Public Function Load(CatalogueID As Long) As Boolean
    On Error GoTo errHandler
Dim objPersist As a_PT_P
Dim tmp As String

    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    ' code to load the object goes here
    Set objPersist = New a_PT_P
    Load = True
    tmp = objPersist.Fetch(CatalogueID)
    If tmp > "" Then
        SetSuperState tmp
    Else
        Load = False
    End If
    Set objPersist = Nothing
    mobjValid.RuleBroken "CODE", False
    
    
    If mudtProps.B1Min < 0 Then mudtProps.B1Min = 0
    If mudtProps.B1Max < mudtProps.B1Min Then mudtProps.B1Max = mudtProps.B1Min
    
    If mudtProps.B2Min < mudtProps.B1Max Then mudtProps.B2Min = mudtProps.B1Max
    If mudtProps.B2Max < mudtProps.B2Min Then mudtProps.B2Max = mudtProps.B2Min
    
    If mudtProps.B3Min < mudtProps.B2Max Then mudtProps.B3Min = mudtProps.B2Max
    If mudtProps.B3Max < mudtProps.B3Min Then mudtProps.B3Max = mudtProps.B3Min
    
    SetDirty True
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Load(CatalogueID)", CatalogueID
End Function

Public Property Get PTID() As Long
    On Error GoTo errHandler
    PTID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.PTID"
End Property
Public Function SetCRSALES(str As String)
    mudtProps.CRSALES = FNS(str)
End Function
Public Function SetCRSALES_CONTRA(str As String)
    mudtProps.CRSALES_CONTRA = FNS(str)
End Function
Public Function SetCASHSALES(str As String)
    mudtProps.CASHSALES = FNS(str)
End Function
Public Function SetCASHSALES_CONTRA(str As String)
    mudtProps.CASHSALES_CONTRA = FNS(str)
End Function
Public Function SetPURCHASES(str As String)
    mudtProps.PURCHASES = FNS(str)
End Function
Public Function SetPURCHASES_CONTRA(str As String)
    mudtProps.PURCHASES_CONTRA = FNS(str)
End Function
Public Function SetVAT(str As String)
    mudtProps.VAT = FNS(str)
End Function

Public Property Get Number() As String
    On Error GoTo errHandler
  Number = FNS(mudtProps.Number)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Number"
End Property
Public Function SetNumber(val As String) As Boolean
    If FNS(mudtProps.SystemCode) > "" Then
        SetNumber = True
        Exit Function
    End If
    If mcolStack.Count = 0 Then Err.Raise 383
    SetNumber = True
    If Len(val) > Len(mudtProps.Number) Then
        Err.Raise 384
    End If
    mudtProps.Number = val
    SetNumber = ValidateObject("Number")
    mudtProps.IsDirty = True
    Exit Function
End Function



Public Property Get code() As String
    On Error GoTo errHandler
  code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Code"
End Property
Public Function SetCode(val As String) As Boolean
    If FNS(mudtProps.SystemCode) > "" Then
        SetCode = True
        Exit Function
    End If
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCode = True
    If Len(val) > Len(mudtProps.code) Then
        Err.Raise 384
    End If
    mudtProps.code = val
    SetCode = ValidateObject("CODE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Discount
    bOK = SetField_DOUBLE(mudtProps.Discount, val, "DISCOUNT", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Discount) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetDiscount(val)", val
End Function

Public Function SetRound(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Round
    bOK = SetField_LONG(mudtProps.Round, val, "ROUND", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Round) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("ROUND")
    End If
    SetRound = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetRound(val)", val
End Function
Public Function SetB1Min(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B1Min
    bOK = SetField_LONG(mudtProps.B1Min, val, "B1", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B1Min) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B1")
    End If
    SetB1Min = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB1Min(val)", val
End Function
Public Function SetB1Max(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B1Max
    bOK = SetField_LONG(mudtProps.B1Max, val, "B1", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B1Max) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B1")
    End If
    SetB1Max = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB1Max(val)", val
End Function
Public Function SetB1MU(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B1MU
    bOK = SetField_LONG(mudtProps.B1MU, val, "B1", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B1MU) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B1")
    End If
    SetB1MU = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB1MU(val)", val
End Function

Public Function SetB2Min(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B2Min
    bOK = SetField_LONG(mudtProps.B2Min, val, "B2", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B2Min) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B2")
    End If
    SetB2Min = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB2Min(val)", val
End Function
Public Function SetB2Max(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B2Max
    bOK = SetField_LONG(mudtProps.B2Max, val, "B2", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B2Max) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B2")
    End If
    SetB2Max = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB2Max(val)", val
End Function
Public Function SetB2MU(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B2MU
    bOK = SetField_LONG(mudtProps.B2MU, val, "B2", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B2MU) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B2")
    End If
    SetB2MU = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB2MU(val)", val
End Function

Public Function SetB3Min(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B3Min
    bOK = SetField_LONG(mudtProps.B3Min, val, "B3", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B3Min) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B3")
    End If
    SetB3Min = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB3Min(val)", val
End Function
Public Function SetB3Max(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B3Max
    bOK = SetField_LONG(mudtProps.B3Max, val, "B3", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B3Max) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B3")
    End If
    SetB3Max = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB3Max(val)", val
End Function
Public Function SetB3MU(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.B3MU
    bOK = SetField_LONG(mudtProps.B3MU, val, "B3", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.B3MU) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("B3")
    End If
    SetB3MU = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetB3MU(val)", val
End Function
Public Property Let SaleOrReturn(val As Boolean)
    On Error GoTo errHandler
    mudtProps.SaleOrReturn = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SaleOrReturn(val)", val
End Property
Public Property Get RoundedS() As String
    On Error GoTo errHandler
    RoundedS = CStr(mudtProps.Round)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.RoundedS"
End Property
Public Property Get Rounded() As Long
    On Error GoTo errHandler
    Rounded = mudtProps.Round
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Rounded"
End Property

Public Property Get B1MinS() As String
    On Error GoTo errHandler
    B1MinS = CStr(mudtProps.B1Min)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B1MinS"
End Property
Public Property Get B1Min() As Long
    On Error GoTo errHandler
    B1Min = mudtProps.B1Min
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B1Min"
End Property

Public Property Get B1MaxS() As String
    On Error GoTo errHandler
    B1MaxS = CStr(mudtProps.B1Max)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B1MaxS"
End Property
Public Property Get B1Max() As Long
    On Error GoTo errHandler
    B1Max = mudtProps.B1Max
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B1Max"
End Property


Public Property Get B1MUS() As String
    On Error GoTo errHandler
    B1MUS = CStr(mudtProps.B1MU)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B1MUS"
End Property
Public Property Get B1MU() As Long
    On Error GoTo errHandler
    B1MU = mudtProps.B1MU
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B1MU"
End Property



Public Property Get B2MinS() As String
    On Error GoTo errHandler
    B2MinS = CStr(mudtProps.B2Min)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B2MinS"
End Property
Public Property Get B2Min() As Long
    On Error GoTo errHandler
    B2Min = mudtProps.B2Min
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B2Min"
End Property


Public Property Get B2MaxS() As String
    On Error GoTo errHandler
    B2MaxS = CStr(mudtProps.B2Max)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B2MaxS"
End Property
Public Property Get B2Max() As Long
    On Error GoTo errHandler
    B2Max = mudtProps.B2Max
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B2Max"
End Property


Public Property Get B2MUS() As String
    On Error GoTo errHandler
    B2MUS = CStr(mudtProps.B2MU)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B2MUS"
End Property
Public Property Get B2MU() As Long
    On Error GoTo errHandler
    B2MU = mudtProps.B2MU
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B2MU"
End Property

Public Property Get B3MinS() As String
    On Error GoTo errHandler
    B3MinS = CStr(mudtProps.B3Min)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B3MinS"
End Property
Public Property Get B3Min() As Long
    On Error GoTo errHandler
    B3Min = mudtProps.B3Min
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B3Min"
End Property

Public Property Get B3MaxS() As String
    On Error GoTo errHandler
    B3MaxS = CStr(mudtProps.B3Max)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B3MaxS"
End Property
Public Property Get B3Max() As Long
    On Error GoTo errHandler
    B3Max = mudtProps.B3Max
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B3Max"
End Property


Public Property Get B3MUS() As String
    On Error GoTo errHandler
    B3MUS = CStr(mudtProps.B3MU)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B3MUS"
End Property
Public Property Get B3MU() As Long
    On Error GoTo errHandler
    B3MU = mudtProps.B3MU
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.B3MU"
End Property


Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.Discount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Discount"
End Property
Public Property Get DiscountF() As String
    On Error GoTo errHandler
    DiscountF = PBKSPercentF(Round(mudtProps.Discount, 2))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.DiscountF"
End Property
Public Property Get SaleOrReturn() As Boolean
    On Error GoTo errHandler
    SaleOrReturn = mudtProps.SaleOrReturn
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SaleOrReturn"
End Property
Public Property Get IsVoucher() As Boolean
    On Error GoTo errHandler
    IsVoucher = (Trim(mudtProps.SystemCode) = "V1")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.IsVoucher"
End Property
Public Property Let SetVoucher(val As Boolean)
    On Error GoTo errHandler
    mudtProps.IsVoucher = FNB(val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetVoucher(val)", val
End Property
Public Property Get ActiveYN() As Boolean
    On Error GoTo errHandler
    ActiveYN = mudtProps.Active
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.ActiveYN"
End Property
Public Property Let ActiveYN(val As Boolean)
    On Error GoTo errHandler
    If FNS(mudtProps.SystemCode) > "" Then
        mudtProps.Active = True
        mudtProps.IsDirty = True
        Exit Property
    End If
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Active = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.ActiveYN(val)", val
End Property



Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.IsNew"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.IsEditing"
End Property
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.IsValid"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.IsDirty"
End Property
Public Function Validate(pMsg As String) As Boolean
    On Error GoTo errHandler
    pMsg = mobjValid.AllBrokenRules
    Validate = (pMsg = "")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.Validate(pMsg)", pMsg
End Function
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If Not FNS(mudtProps.code) > "" Then
        mobjValid.BreakRule "CODE", True
        If pFld = "CODE" Then bValid = False
    Else
        mobjValid.BreakRule "CODE", False
    End If
    If (mudtProps.B1Max < mudtProps.B1Min) Or mudtProps.B1MU < 0 Then
        mobjValid.BreakRule "B1", True
    Else
        mobjValid.BreakRule "B1", False
    End If
    If (mudtProps.B2Max < mudtProps.B2Min) Or mudtProps.B2MU < 0 Then
        mobjValid.BreakRule "B2", True
    Else
        mobjValid.BreakRule "B2", False
    End If
    If (mudtProps.B3Max < mudtProps.B3Min) Or mudtProps.B3MU < 0 Then
        mobjValid.BreakRule "B3", True
    Else
        mobjValid.BreakRule "B3", False
    End If
    If mudtProps.Discount < -100 Or mudtProps.Discount > 100 Then
        mobjValid.BreakRule "DISCOUNT", True
    Else
        mobjValid.BreakRule "DISCOUNT", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.ValidateObject(pFld)", pFld
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
    ErrorIn "a_PT.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing code", "CODE"
    colClassors.Add "or in band 1", "B1"
    colClassors.Add "or in band 2", "B2"
    colClassors.Add "or in band 3", "B3"
    colClassors.Add "Invalid discount", "DISCOUNT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.LoadClassorsCollection"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PT.SetDirty(pVal)", pVal
End Sub

Public Property Get CRSALES() As String
    CRSALES = FNS(mudtProps.CRSALES)
    SetDirty True
End Property
Public Property Get CRSALES_CONTRA() As String
    CRSALES_CONTRA = FNS(mudtProps.CRSALES_CONTRA)
    SetDirty True
End Property
Public Property Get CASHSALES() As String
    CASHSALES = FNS(mudtProps.CASHSALES)
    SetDirty True
End Property
Public Property Get CASHSALES_CONTRA() As String
    CASHSALES_CONTRA = FNS(mudtProps.CASHSALES_CONTRA)
    SetDirty True
End Property
Public Property Get PURCHASES() As String
    PURCHASES = FNS(mudtProps.PURCHASES)
    SetDirty True
End Property
Public Property Get PURCHASES_CONTRA() As String
    PURCHASES_CONTRA = FNS(mudtProps.PURCHASES_CONTRA)
    SetDirty True
End Property
Public Property Get VAT() As String
    VAT = FNS(mudtProps.VAT)
    SetDirty True
End Property

