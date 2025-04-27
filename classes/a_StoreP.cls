VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_StoreP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As StorePProps
Private mcolStack As Collection
Private mobjParent As ch_StoreP
Event Valid(IsValid As Boolean)
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Public Sub BeginEdit()
    On Error GoTo errHandler
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.IsEditing"
End Property

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddStoreP Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.ApplyEdit"
End Sub
Friend Sub Initialize(Parent As ch_StoreP)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
    LoadClassorsCollection
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
    
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.IsValid"
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.mobjValid_Status(pMsg)", pMsg
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As StorePData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As StorePData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.SetSuperState(buffer)", buffer
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.Delete"
End Sub


Public Property Get STID() As Long
    On Error GoTo errHandler
    STID = mudtProps.STID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.STID"
End Property
Public Property Get StoreName() As String
    On Error GoTo errHandler
    StoreName = FNS(mudtProps.StoreName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.StoreName"
End Property
Public Property Get FirstReceived() As Date
    On Error GoTo errHandler
    FirstReceived = FND(mudtProps.FirstReceivedDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.FirstReceived"
End Property
Public Property Get FirstReceivedF() As Date
    On Error GoTo errHandler
    FirstReceivedF = Format(FND(mudtProps.FirstReceivedDate), "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.FirstReceivedF"
End Property
Public Property Get LastReceived() As Date
    On Error GoTo errHandler
    LastReceived = FND(mudtProps.LastDeliveredDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastReceived"
End Property
Public Property Get LastReceivedF() As String
    On Error GoTo errHandler
    
    If FND(mudtProps.LastDeliveredDate) < "2000-01-01" Then
        LastReceivedF = ""
    Else
        LastReceivedF = Format(FND(mudtProps.LastDeliveredDate), "dd/mm/yyyy")
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastReceivedF"
End Property
Public Property Get LastSoldDate() As Date
    On Error GoTo errHandler
    LastSoldDate = FND(mudtProps.LastSoldDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastSoldDate"
End Property
Public Property Get LastSoldDateF() As String
    On Error GoTo errHandler
    If FND(mudtProps.LastSoldDate) < "2000-01-01" Then
        LastSoldDateF = ""
    Else
        LastSoldDateF = Format(FND(mudtProps.LastSoldDate), "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastSoldDateF"
End Property

'Public Property Get LastSharedDate() As Date
'    On Error GoTo ErrHandler
'    LastSharedDate = FND(mudtProps.LastSharedDate)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_StoreP.LastSharedDate"
'End Property
'Public Property Get LastSharedDateF() As String
'    On Error GoTo ErrHandler
'    LastSharedDateF = Format(FND(mudtProps.LastSharedDate), "dd/mm/yyyy")
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_StoreP.LastSharedDateF"
'End Property
'Public Property Get LastSharedDateFShort() As String
'    On Error GoTo ErrHandler
'    If CDate(mudtProps.LastSharedDate) = 0 Then
'        LastSharedDateFShort = ""
'    Else
'        LastSharedDateFShort = Format(FND(mudtProps.LastSharedDate), "dd/mm")
'    End If
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_StoreP.LastSharedDateFShort"
'End Property
'Public Property Get LastSharedDateFShortwithParentheses() As String
'    On Error GoTo ErrHandler
'
'    LastSharedDateFShortwithParentheses = "(" & LastSharedDateFShort & ")"
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_StoreP.LastSharedDateFShort"
'End Property



Public Property Get PID() As String
    On Error GoTo errHandler
    PID = Trim(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.PID"
End Property
Public Property Get QtyOnHand() As Long
    On Error GoTo errHandler
    QtyOnHand = FNN(mudtProps.QtyOnHand)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.QtyOnHand"
End Property
Public Property Get QtyCopiesOnHand() As Long
    On Error GoTo errHandler
    QtyCopiesOnHand = FNN(mudtProps.QtyCopiesOnHand)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.QtyCopiesOnHand"
End Property
Public Property Get QtyonOrder() As Long
    On Error GoTo errHandler
    QtyonOrder = FNN(mudtProps.QtyonOrder)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.QtyOnOrder"
End Property
Public Property Get QtyReserved() As Long
    On Error GoTo errHandler
    QtyReserved = FNN(mudtProps.QtyReserved)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.QtyReserved"
End Property
Public Property Get SP() As Double
    On Error GoTo errHandler
    SP = FNDBL(mudtProps.SP)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.SP"
End Property
Public Property Get TotalQtySold() As Long
    On Error GoTo errHandler
    TotalQtySold = FNN(mudtProps.TotalSold)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.TotalQtySold"
End Property

Public Property Get LastDeliveredPrice() As Double
    On Error GoTo errHandler
    LastDeliveredPrice = FNDBL(mudtProps.LastDeliveredPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastDeliveredPrice"
End Property
Public Property Get LastDeliveredPriceF() As String
    On Error GoTo errHandler
    LastDeliveredPriceF = Format(mudtProps.LastDeliveredPrice, "###,##0.00")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastDeliveredPrice"
End Property
Public Property Get QtyOnBackorder() As Long
    On Error GoTo errHandler
    QtyOnBackorder = FNN(mudtProps.QtyOnBackorder)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.QtyOnBackorder"
End Property
Public Property Get DateLastStocktake() As Date
    On Error GoTo errHandler
    DateLastStocktake = FND(mudtProps.DateLastStocktake)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.DateLastStocktake"
End Property
Public Property Get DateLastStocktakeF() As String
    On Error GoTo errHandler
    DateLastStocktakeF = Format(DateLastStocktake, "###,##0.00")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.DateLastStocktakeF"
End Property

Public Property Get LastOrderedDate() As Date
    On Error GoTo errHandler
    LastOrderedDate = FND(mudtProps.LastOrderedDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastOrderedDate"
End Property
Public Property Get LastOrderedDateF() As String
    On Error GoTo errHandler
    LastOrderedDateF = Format(LastOrderedDate, "###,##0.00")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastOrderedDateF"
End Property
Public Property Get LastOrderedDateF2() As String
    On Error GoTo errHandler
    If LastOrderedDate > CDate(0) Then
        LastOrderedDateF2 = Format(LastOrderedDate, "DD/MM/YYYY") & "(" & CStr(Me.LastQtyFirmOrdered) & "/" & CStr(Me.LastQtySSOrdered) & ")"
    Else
        LastOrderedDateF2 = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastOrderedDateF2"
End Property


Public Property Get QtyLastStocktake() As Long
    On Error GoTo errHandler
    QtyLastStocktake = FNN(mudtProps.QtyLastStocktake)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.QtyLastStocktake"
End Property
Public Property Get LastQtyFirmOrdered() As Long
    On Error GoTo errHandler
    LastQtyFirmOrdered = FNN(mudtProps.LastQtyFirmOrdered)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastQtyFirmOrdered"
End Property
Public Property Get LastQtySSOrdered() As Long
    On Error GoTo errHandler
    LastQtySSOrdered = FNN(mudtProps.LastQtySSOrdered)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LastQtySSOrdered"
End Property
'Private Function ValidateObject(pVal As String, pFld As String) As Boolean
'
'    ValidateObject = True
'    If Not (Len(Trim(mudtProps.Fax)) > 3 Or Len(Trim(mudtProps.Fax)) = 0) Then
'        mobjValid.BreakRule "FAX", True
'        If pFld = "FAX" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "FAX", False
'    End If
'    If Not (Len(Trim(mudtProps.Description)) > 2) Then
'        mobjValid.BreakRule "DESCRIPTION", True
'        If pFld = "DESCRIPTION" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "DESCRIPTION", False
'    End If
'    If Not (Len(Trim(mudtProps.Line1)) > 2) Then
'        mobjValid.BreakRule "LINE1", True
'        If pFld = "LINE1" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "LINE1", False
'    End If
'    If Not (Len(Trim(mudtProps.pCode)) > 3 Or Len(Trim(mudtProps.pCode)) = 0) Then
'        mobjValid.BreakRule "PCODE", True
'        If pFld = "PCODE" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "PCODE", False
'    End If
'    If Not (Len(Trim(mudtProps.Phone)) > 5 Or Len(Trim(mudtProps.Phone)) = 0) Then
'        mobjValid.BreakRule "PHONE", True
'        If pFld = "PHONE" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "PHONE", False
'    End If
'    If Not (Len(Trim(mudtProps.Phone)) > 5 Or Len(Trim(mudtProps.Phone)) = 0) Then
'        mobjValid.BreakRule "PHONE", True
'        If pFld = "PHONE" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "PHONE", False
'    End If
'    If Not (Len(Trim(mudtProps.BusPhone)) > 5 Or Len(Trim(mudtProps.BusPhone)) = 0) Then
'        mobjValid.BreakRule "BUSPHONE", True
'        If pFld = "BUSPHONE" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "BUSPHONE", False
'    End If
'    If Not (Len(Trim(mudtProps.Fax)) > 5 Or Len(Trim(mudtProps.Fax)) = 0) Then
'        mobjValid.BreakRule "FAX", True
'        If pFld = "FAX" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "FAX", False
'    End If
'    If Not (Len(Trim(mudtProps.Email)) > 5 Or Len(Trim(mudtProps.Email)) = 0) Then
'        mobjValid.BreakRule "EMAIL", True
'        If pFld = "EMAIL" Then ValidateObject = False
'    Else
'        mobjValid.BreakRule "EMAIL", False
'    End If
'    mobjValid.GetStatus
'End Function
'Private Function TranslateErrors(pRawors As String) As String
'Dim strOut As String
'Dim strRule, strAllRules As String
'Dim NoMoreRules As Boolean
'Dim iMarker, iStart As Integer
'    iMarker = 1
'    strAllRules = ""
'    If Len(pRawors) > 0 Then
'        iMarker = InStr(iMarker + 1, pRawors, ",")
'        If iMarker > 0 Then
'            strAllRules = colClassors(Left(pRawors, iMarker - 1))
'        Else
'            strAllRules = colClassors(pRawors)
'        End If
'        Do Until iMarker = 0
'            iStart = iMarker + 1
'            iMarker = InStr(iStart, pRawors, ",")
'            If iMarker > 0 Then
'                strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
'            Else
'                strRule = colClassors(Mid(pRawors, iStart))
'            End If
'
'            strAllRules = strAllRules & vbCrLf & strRule
'        Loop
'    End If
'    TranslateErrors = strAllRules
'End Function
'Private Sub LoadClassorsCollection()
'    Set colClassors = New Collection
'    colClassors.Add "Address description too short", "DESCRIPTION"
'    colClassors.Add "Address line one too short", "LINE1"
'    colClassors.Add "Addressee too short", "ADDRESSEE"
'    colClassors.Add "Phone too short", "PHONE"
'    colClassors.Add "Business phone too short", "BUSPHONE"
'    colClassors.Add "Fax too short", "FAX"
'    colClassors.Add "Email too short", "EMAIL"
'End Sub
'
'
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
    ErrorIn "a_StoreP.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.LoadClassorsCollection"
End Sub

Public Property Get StoreCode() As String
    On Error GoTo errHandler
    StoreCode = FNS(mudtProps.StoreCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP.StoreCode"
End Property
