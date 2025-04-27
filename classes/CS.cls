VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngLPLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLPExt_tot As Long
Dim lngFPExt_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Private tlOperators As z_TextList
Private mudtProps As CSProps
Private mcolStack As Collection
Private mColCSLines As ch_CSL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Event Valid(pMsg As String)
Event TotalChange(lngTotal As String, lngTotalForeign As String)

Public Function CalculateTotal()
    On Error GoTo errHandler
Dim ol As a_CSL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    For Each ol In mColCSLines
            ol.RecalculateLine
            lngLPExt_tot = lngLPExt_tot + ol.PExt
            lngVAT_tot = lngVAT_tot + ol.PLessDiscExtVAT
            lngDisc_tot = lngDisc_tot + ol.PDiscExt
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + ol.PLessDiscExt
    Next
    mudtProps.TotalExtensionSimple = lngLPExt_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalPayable = lngLPLessDiscExt_tot ' - lngLDeposit_tot
    RaiseEvent TotalChange(Me.TotalLessDiscExtF, Me.TotalLessDiscExtF)
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.CalculateTotal"
End Function
Public Property Get TotalLessDiscExtF() As String
    On Error GoTo errHandler
        TotalLessDiscExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalLessDiscExtF"
End Property
Public Property Get TotalExtF() As String
    On Error GoTo errHandler
        TotalExtF = Format(mudtProps.TotalExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalExtF"
End Property

Public Property Get TotalLessDiscExt() As String
    On Error GoTo errHandler
        TotalLessDiscExt = mudtProps.TotalExtension
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalLessDiscExt"
End Property
Public Property Get TotalVATF() As String
    On Error GoTo errHandler
        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalVATF"
End Property
Public Property Get TotalPayable() As String
    On Error GoTo errHandler
        TotalPayable = mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalPayable"
End Property
Public Property Get TotalPayableF() As String
    On Error GoTo errHandler
        TotalPayableF = Format(TotalPayable, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalPayableF"
End Property
Public Property Get TotalDiscount() As Long
    On Error GoTo errHandler
        TotalDiscount = mudtProps.TotalDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalDiscount"
End Property
Public Property Get TotalDiscountF() As String
    On Error GoTo errHandler
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TotalDiscountF"
End Property
Public Function DisplayTotals(pLabels As String, pValues As String) 'This must be the same as PRINTINVOICE
    On Error GoTo errHandler
Dim strLabels As String
Dim strValues As String
    strLabels = ""
    strValues = ""
    If oPC.Configuration.IsVATRegion Then
            strLabels = strLabels & Chr(13) & "Total"
            strValues = strValues & Chr(13) & Me.TotalExtF
            strLabels = strLabels & Chr(13) & "Less discount"
            strValues = strValues & Chr(13) & Me.TotalDiscountF
            strLabels = strLabels & Chr(13) & "Includes VAT of "
            strValues = strValues & Chr(13) & Me.TotalVATF
    Else
            strLabels = strLabels & Chr(13) & "subtotal"
            strValues = strValues & Chr(13) & Me.TotalExtF
            strLabels = strLabels & Chr(13) & "Less discount"
            strValues = strValues & Chr(13) & Me.TotalDiscountF
            strLabels = strLabels & Chr(13) & "Total"
            strValues = strValues & Chr(13) & Me.TotalPayableF
    End If
    pLabels = strLabels
    pValues = strValues
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DisplayTotals(pLabels,pValues)", Array(pLabels, pValues)
End Function
Public Sub VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    If Me.Status <> stVOID Then
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.VoidDocument"
End Sub
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.Status"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set mColCSLines = New ch_CSL
    mColCSLines.component Me
    Set mcolStack = New Collection
    LoadClassorsCollection
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    mudtProps.Status = 2
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set mobjValid = Nothing
    Set mColCSLines = Nothing
    Set mcolStack = Nothing
    Set colClassors = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
  RaiseEvent Valid(oRS)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.mobjValid_BrokenRule(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
  RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.mobjValid_NoBrokenRules"
End Sub
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As CSData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As CSData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "CSLS", mColCSLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.GetSuperState"
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
    mColCSLines.SetSuperState .ReadProperty("CSLS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.SetSuperState(buffer)", buffer
End Sub

Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.StatusF"
End Property

Public Property Let Status(val As Long)
    On Error GoTo errHandler
    mudtProps.Status = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.Status(val)", val
End Property
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.SetStatus(val)", val
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
  mColCSLines.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mColCSLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
Dim objPersist As a_CS_P
  If mcolStack.Count = 0 Then Err.Raise 445
    If DOCCode = "" Then
        DOCCode = GetCSNumber("GETANDWRITE")
    End If
    If CaptureDate < #1/1/1995# And mudtProps.IsNew Then CaptureDate = Now

  mColCSLines.ApplyEdit
  Set objPersist = New a_CS_P
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject mudtProps.TRID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    mcolStack.Remove mcolStack.Count
    SetSuperState objPersist.Save(GetSuperState)
    mudtProps.IsNew = False
  Else
    mcolStack.Remove mcolStack.Count
  End If
  Set objPersist = Nothing
  mudtProps.IsDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.ApplyEdit"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mColCSLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.IsValid"
End Property
Public Function GetStatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.GetSTatus"
End Function

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.Delete"
End Sub
Public Sub Load(pCSID As Long)
    On Error GoTo errHandler
Dim oPersist As a_CS_P

    If pCSID = 0 Then Exit Sub
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    Set oPersist = New a_CS_P
    SetSuperState oPersist.Fetch(pCSID)
    Set oPersist = Nothing
    CalculateTotal
    mudtProps.IsNew = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.Load(pCSID)", pCSID
End Sub
Public Sub LoadExistingTA(lngResult As Long, pDate As Date, pComputer As String)
    On Error GoTo errHandler
Dim oPersist As a_CS_P
Dim str As String

    Set oPersist = New a_CS_P
    str = oPersist.FetchOpenTAForThisWorkstation(lngResult, pDate, Me.TILLID)
    Set oPersist = Nothing
    If lngResult = 0 Then
        mudtProps.IsNew = False
        SetSuperState str
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.LoadExistingTA(lngResult,pDate,pComputer)", Array(lngResult, pDate, pComputer)
End Sub
'Public Sub Delete()
'    If Not flgEditing Then Err.Raise 445
'    flgDeleted = True
'End Sub
Private Sub DeleteObject(lngResult As Long, ID As Long)
    On Error GoTo errHandler

  Dim strSQL As String
  Dim lngRecs As Long
  
    strSQL = "DELETE * FROM tblTA WHERE Transaction_ID=" & ID
    oPC.COShort.execute strSQL
    
EXIT_Handler:
'Err_Handler:
'    If (( = DB_EXISTS) Or ( = DB_MISSING) Or ( = DB_REF) Or ( = DB_FIELDCONSTRAINT)) Then
'        lngResult =
'    Else
'        tmp =
'        tmpor = or
'        oor.Setor tmp, tmpor, Now(), "a_CS", "DeleteObject", "unknown"
'        Err.Raise tmp
'    End If
'    GoTo EXIT_Handler
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DeleteObject(lngResult,ID)", Array(lngResult, ID)
End Sub
Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TRID"
End Property
Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TPID"
End Property
Public Property Let TPID(val As Long)
    On Error GoTo errHandler
    If Len(val) > Len(mudtProps.TPID) Then Err.Raise vbObjectError + 1001, "String value too long"
    mobjValid.RuleBroken "TPID", (Len(Trim$(val)) = 0)
    mudtProps.TPID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TPID(val)", val
End Property
Public Property Get TILLID() As Long
    On Error GoTo errHandler
    TILLID = mudtProps.TILLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TILLID"
End Property
Public Property Let TILLID(val As Long)
    On Error GoTo errHandler
    mudtProps.TILLID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TILLID(val)", val
End Property
Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DocCode"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DocCode(val)", val
End Property
'Public Property Get TextFileFullPathAndName() As String
'    TextFileFullPathAndName = Trim$(mudtProps.TextFileFullPathAndName)
'End Property
Public Property Let TextFileFullPathAndName(val As String)
    On Error GoTo errHandler
    If Len(val) > Len(mudtProps.TextFileFullPathAndName) Then Err.Raise vbObjectError + 1001, "String value too long"
    mobjValid.RuleBroken "TextFileFullPathAndName", (Len(Trim$(val)) = 0)
    mudtProps.TextFileFullPathAndName = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TextFileFullPathAndName(val)", val
End Property
Public Property Get FrontDeskComputerName() As String
    On Error GoTo errHandler
    FrontDeskComputerName = Trim$(mudtProps.FrontDeskComputerName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.FrontDeskComputerName"
End Property
Public Property Let FrontDeskComputerName(val As String)
    On Error GoTo errHandler
    If Len(val) > Len(mudtProps.FrontDeskComputerName) Then Err.Raise vbObjectError + 1001, "String value too long"
    mobjValid.RuleBroken "FrontDeskComputerName", (Len(Trim$(val)) = 0)
    mudtProps.FrontDeskComputerName = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.FrontDeskComputerName(val)", val
End Property
'Public Property Get Note() As String
'    Note = Trim$(mudtProps.Note)
'End Property
'Public Property Let Note(val As String)
'    If Len(val) > Len(mudtProps.Note) Then Err.Raise vbObjectError + 1001, "String value too long"
''    mobjValid.RuleBroken "Note", (Len(Trim$(val)) = 0)
'    mudtProps.Note = val
'    SetDirty True
'End Property
'Public Property Get OperatorID() As Long
'    OperatorID = udtProps.OperatorID
'End Property
'Public Property Let OperatorID(val As Long)
'    mobjValid.RuleBroken "Operator", (val = 0)
'    mudtProps.OperatorID = val
'    SetDirty True
'End Property
Public Property Get TAType() As Long
    On Error GoTo errHandler
    TAType = mudtProps.TAType
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TAType"
End Property
Public Property Let TAType(val As Long)
    On Error GoTo errHandler
    mobjValid.RuleBroken "TAType", (val = 0)
    mudtProps.TAType = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.TAType(val)", val
End Property
Public Property Let DateIssued(val As Date)
    On Error GoTo errHandler
    
    mobjValid.RuleBroken "DateIssued", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.DateIssued = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DateIssued(val)", val
End Property
Public Property Get DateIssued() As Date
    On Error GoTo errHandler
    DateIssued = mudtProps.DateIssued
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DateIssued"
End Property
Public Property Get DateIssuedF() As String
    On Error GoTo errHandler
    If CLng(mudtProps.DateIssued) = 0 Then
        DateIssuedF = ""
    Else
        DateIssuedF = Format(mudtProps.DateIssued, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DateIssuedF"
End Property
Public Property Let CaptureDate(val As Date)
    On Error GoTo errHandler
    
    mobjValid.RuleBroken "CaptureDate", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.CaptureDate(val)", val
End Property
Public Property Get CaptureDate() As Date
    On Error GoTo errHandler
    CaptureDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.CaptureDate"
End Property
Public Property Get CaptureDateF() As String
    On Error GoTo errHandler
    If CLng(mudtProps.CaptureDate) = 0 Then
        CaptureDateF = ""
    Else
        CaptureDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.CaptureDateF"
End Property
Public Property Let DateStarted(val As String)
    On Error GoTo errHandler
    
    mobjValid.RuleBroken "DateStarted", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.DateStarted = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DateStarted(val)", val
End Property
Public Property Get DateStarted() As String
    On Error GoTo errHandler
    If CLng(mudtProps.DateStarted) = 0 Then
        DateStarted = ""
    Else
        DateStarted = Trim$(mudtProps.DateStarted)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.DateStarted"
End Property
Public Property Get Operators() As z_TextList
    Set Operators = tlOperators
End Property
'Public Sub Issue(lngResult As Long)
'On Error GoTo err_Handler
'
'
'    BeginEdit
'        udtProps.Status = 3
'        udtProps.DateIssued = Now()
'        flgDirty = True
'    ApplyEdit lngResult, False
'
'
'EXIT_Handler:
'    Exit Sub
'Err_Handler:
'    tmp =
'    tmpor = or
'    oor.Setor tmp, tmpor, Now(), "a_CS", "Issue", "unknown"
'    Err.Raise tmp
'    GoTo EXIT_Handler
'End Sub

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
    ErrorIn "a_CS.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.LoadClassorsCollection"
End Sub
Public Function GetCSNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strWork As String

    Set oCode = New z_Code
    
    strWork = oCode.GetNextCode(enCashSale, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetCSNumber = strDefaultCo & "C" & strWork
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.GetCSNumber(strTmp)", strTmp
End Function

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.SetDirty(pVal)", pVal
End Sub
Public Property Get CSLines() As ch_CSL
    Set CSLines = mColCSLines
End Property
Public Function SetTILLID(pName As String) As Long
    On Error GoTo errHandler
Dim oP As a_CS_P
Dim lngTILLID As Long

    Set oP = New a_CS_P
    lngTILLID = oP.GetTILLID(pName)
    Me.TILLID = lngTILLID
    Set oP = Nothing
    SetTILLID = lngTILLID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS.SetTILLID(pName)", pName
End Function

