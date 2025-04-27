VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Dim lngLPLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLPLessDiscExtNonVAT_tot As Long
Dim lngLPExt_tot As Long
Dim lngFPExt_tot As Long
Dim lngVAT_tot As Long
Dim lngTotalQty As Long
Dim lngDisc_tot As Long
Private tlOperators As z_TextList
Private mudtProps As RProps
Private mcolStack As Collection
Private mCOLRLines As ch_RL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Event Valid(pMsg As String)
Event TotalChange(lngTotal As String, lngTotalForeign As String)
Event Dirty(pVal As Boolean)
Private WithEvents mSupplier As a_Supplier
Attribute mSupplier.VB_VarHelpID = -1
Private xMLDoc As ujXML

Const CDOCCODE = "R"
Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim ol As a_RL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngTotalQty = 0
    lngLPLessDiscExtNonVAT_tot = 0
    For Each ol In mCOLRLines
        If ol.Status <> "CAN" Then
            ol.RecalculateLine
            If Me.Status = stISSUED Then
                lngTotalQty = lngTotalQty + ol.QtyRequested
            ElseIf Me.Status = stCOMPLETE Then
                lngTotalQty = lngTotalQty + (ol.QtyReturned - ol.QtyRejected)
            End If
            lngLPExt_tot = lngLPExt_tot + ol.PExt(False)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(False)
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + ol.PLessDiscExt(False)
        End If
    Next
    mudtProps.TotalExtensionSimple = lngLPExt_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalQty = lngTotalQty
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngVAT_tot
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim ol As a_RL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngFPLessDiscExt_tot = 0
    lngFPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngTotalQty = 0
    For Each ol In mCOLRLines
        If ol.Status <> "CAN" Then
            ol.RecalculateLine
            If Me.Status = stISSUED Then
                lngTotalQty = lngTotalQty + ol.QtyRequested
            ElseIf Me.Status = stCOMPLETE Then
                lngTotalQty = lngTotalQty + (ol.QtyReturned - ol.QtyRejected)
            End If
            lngFPExt_tot = lngFPExt_tot + ol.PExt(True)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(True)
            lngFPLessDiscExt_tot = lngFPLessDiscExt_tot + ol.PLessDiscExt(True)
        End If
    Next
    mudtProps.TotalExtensionSimple_Foreign = lngFPExt_tot
    mudtProps.TotalDiscount_Foreign = lngDisc_tot
    mudtProps.TotalQty = lngTotalQty
    mudtProps.TotalExtension_Foreign = lngFPLessDiscExt_tot
    mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot - lngVAT_tot
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.CalculateTotal_Foreign"
End Function
Public Property Get TotPLessDiscExtNonVAT() As Long
    TotPLessDiscExtNonVAT = lngLPLessDiscExtNonVAT_tot
End Property
Public Property Get TotalLessDiscExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalLessDiscExtF = Format(mudtProps.TotalExtension_Foreign / Supplier.DefaultCurrency.Divisor, Supplier.DefaultCurrency.FormatString)
    Else
        TotalLessDiscExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalLessDiscExtF(bForeign)", bForeign
End Property
Public Property Get TotalExtF() As String
    On Error GoTo errHandler
        TotalExtF = Format(mudtProps.TotalExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalExtF"
End Property
Public Property Get TotalExt() As Long
    On Error GoTo errHandler
        TotalExt = mudtProps.TotalExtensionSimple
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalExt"
End Property
Public Property Get TotalQty() As Long
    On Error GoTo errHandler
        TotalQty = mudtProps.TotalQty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalQty"
End Property


Public Property Get TotalLessDiscExt() As String
    On Error GoTo errHandler
        TotalLessDiscExt = mudtProps.TotalExtension
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalLessDiscExt"
End Property
Public Property Get TotalVATF() As String
    On Error GoTo errHandler
        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalVATF"
End Property
Public Property Get TotalPayable(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayable = mudtProps.TotalPayable_Foreign / CaptureCurrency.Divisor
    Else
        TotalPayable = mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalPayable(bForeign)", bForeign
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayableF = Format(TotalPayable(bForeign), CaptureCurrency.FormatString) & " " & CaptureCurrency.SYSNAME
    Else
        TotalPayableF = Format(TotalPayable(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalPayableF(bForeign)", bForeign
End Property
Public Property Get TotalDiscount() As Long
    On Error GoTo errHandler
        TotalDiscount = mudtProps.TotalDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalDiscount"
End Property
Public Property Get TotalDiscountF() As String
    On Error GoTo errHandler
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TotalDiscountF"
End Property


Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Status"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    mOriginalStatus = stInProcess
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set mCOLRLines = New ch_RL
    mCOLRLines.component Me
    Set mcolStack = New Collection
    LoadClassorsCollection
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    mudtProps.Status = 2
    mudtProps.RType = "S"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set mobjValid = Nothing
    Set mCOLRLines = Nothing
    Set mcolStack = Nothing
    Set colClassors = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
  RaiseEvent Valid(oRS)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.mobjValid_BrokenRule(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
  RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.mobjValid_NoBrokenRules"
End Sub
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As RData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As RData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "RLS", mCOLRLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.GetSuperState"
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
    mCOLRLines.SetSuperState .ReadProperty("RLS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.SetSuperState(buffer)", buffer
End Sub

Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.Status = 1 Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = 2 Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = 3 Then
        StatusF = "AUTHORIZATION REQUESTED"
    ElseIf mudtProps.Status = 4 Then
        StatusF = "STOCK RETURNED"
    ElseIf mudtProps.Status = 7 Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = 6 Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.StatusF"
End Property

Public Property Let Status(val As Long)
    On Error GoTo errHandler
    mudtProps.Status = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Status(val)", val
End Property
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.SetStatus(val)", val
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
  mCOLRLines.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mCOLRLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
Dim objPersist As a_R_P
  If mcolStack.Count = 0 Then Err.Raise 445
    If DOCCode = "" Then
        DOCCode = GetRNumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now

  mCOLRLines.ApplyEdit
  Set objPersist = New a_R_P
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
    ErrorIn "a_R.ApplyEdit"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mCOLRLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IsValid"
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
    ErrorIn "a_R.GetSTatus"
End Function

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Delete"
End Sub
Public Sub Load(pTRID As Long)
    On Error GoTo errHandler
Dim oPersist As a_R_P
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    Set oPersist = New a_R_P
    SetSuperState oPersist.Fetch(pTRID)
    Set oPersist = Nothing
    mOriginalStatus = Me.Status
    Set mSupplier = New a_Supplier
    mSupplier.Load mudtProps.TPID
 '   Set oCurr = mSupplier.DefaultCurrency
    CalculateTotals
    mudtProps.IsNew = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Load(pTRID)", pTRID
End Sub
Public Sub LoadSupplierFromID(val As Long)
    Set mSupplier = Nothing
    Set mSupplier = New a_Supplier
    mSupplier.Load val
End Sub
Public Property Get Supplier() As a_Supplier
    Set Supplier = mSupplier
End Property
Public Property Get RType() As String
     RType = FNS(mudtProps.RType)
End Property

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TRID"
End Property
Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TPID"
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
    ErrorIn "a_R.TPID(val)", val
End Property

Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.StaffID(val)", val
End Property
Public Property Get StaffName() As String
    On Error GoTo errHandler
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).StaffName
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.StaffName"
End Property

Public Property Get StaffNameB() As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
        StaffNameB = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    If strT > "" Then
        StaffNameB = "   (" & strT & ")"
    Else
        StaffNameB = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.StaffNameB"
End Property

Public Property Get TPNAME() As String
    On Error GoTo errHandler
    TPNAME = FNS(mudtProps.TPNAME)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TPName"
End Property
Public Property Let TPNAME(val As String)
    On Error GoTo errHandler
    mudtProps.TPNAME = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.TPName(val)", val
End Property
Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.DocCode"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.DocCode(val)", val
End Property

Public Property Get ApprovalRef() As String
    On Error GoTo errHandler
    ApprovalRef = FNS(mudtProps.ApprovalRef)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ApprovalRef"
End Property
Public Property Let ApprovalRef(val As String)
    mudtProps.ApprovalRef = val
    If Len(val) > Len(mudtProps.ApprovalRef) Then
        Err.Raise 384
    End If
    SetDirty True
    Exit Property
End Property
Public Property Get ApprovalTermDate() As Date
    On Error GoTo errHandler
    ApprovalTermDate = FNS(mudtProps.ApprovalTermDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ApprovalRef"
End Property
Public Property Get ApprovalTermDateF() As String
    On Error GoTo errHandler
    ApprovalTermDateF = Format(ApprovalTermDate, "DD/MM/YYYY")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ApprovalTermDateF"
End Property

Public Property Let ApprovalTermDate(val As Date)
    mudtProps.ApprovalTermDate = val
    SetDirty True
End Property
Public Function SetApprovalTermDate(val As String) As Boolean
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetApprovalTermDate = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetApprovalTermDate = False
        Exit Function
    End If
    
    If dteDate > Date Then
        mobjValid.RuleBroken "APPROVALTERMDATE", False
    Else
        mobjValid.RuleBroken "APPROVALTERMDATE", True
    End If
    
    mudtProps.ApprovalTermDate = dteDate
    SetDirty True

End Function


Public Property Let DOCDate(val As Date)
    On Error GoTo errHandler
    
    mobjValid.RuleBroken "DOCDATE", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.DOCDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.DocDate(val)", val
End Property
Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    If CLng(mudtProps.DOCDate) = 0 Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.DocDateF"
End Property
Public Property Let IssDate(val As Date)
    On Error GoTo errHandler
    
    mobjValid.RuleBroken "CaptureDate", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IssDate(val)", val
End Property
Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IssDate"
End Property
Public Property Get IssDateF() As String
    On Error GoTo errHandler
    If CLng(mudtProps.CaptureDate) = 0 Then
        IssDateF = ""
    Else
        IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.IssDateF"
End Property
Public Property Get Operators() As z_TextList
    Set Operators = tlOperators
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
    ErrorIn "a_R.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing supplier", "TP"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.LoadClassorsCollection"
End Sub
Public Function GetRNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strWork As String

    Set oCode = New z_Code
    
    strWork = oCode.GetNextCode(enReturn, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetRNumber = strDefaultCo & "R" & strWork
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.GetRNumber(strTmp)", strTmp
End Function

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.SetDirty(pVal)", pVal
End Sub
Public Property Get RLines() As ch_RL
    Set RLines = mCOLRLines
End Property

Public Function Post() As String
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    Me.ApplyEdit
    Set oSM = New z_StockManager

    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
            Select Case Status
            Case stVOID
       ''         SetStatus stVOID
            Case stInProcess
                Post = "In Process"
            Case stISSUED
                Post = "Requested"
            Case stCOMPLETE
                oSM.IssueReturn Me.TRID
            Case stPROFORMA
                oSM.IssueReturn Me.TRID
        '        SetStatus stPROFORMA
            End Select
    Case stISSUED
            Select Case Status
            Case stVOID
            Case stISSUED
                Post = "Requested"
            Case stCOMPLETE
                oSM.IssueReturn Me.TRID
            End Select
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    End Select
    Set oSM = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Post"
End Function

'Public Sub PrintTF()
'Dim iTotalLinesCOunt As Integer
'
'    On Error GoTo err_Handler
'    CalculateTotal
'    Set oPCL = New z_PrintClient
'    oPCL.StartRequestPrinting "TFR_" & Me.DocCode
'    With oPC.Configuration.DocumentControls.FindbyTypeName("TRANSFER")
'        oPCL.PrinterSetup .Printer, .PreviewPrint, .Style, .QtyCopies
'    End With
'    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
'        oPCL.WriteTag "STATUS", Me.StatusF
'    End If
'    oPCL.LineCOunt TFLines.Count
'    iTotalLinesCOunt = 3
'    oPCL.LineTotalCount iTotalLinesCOunt 'This must be written before the TOTAL tags for the Printserver requirements
'
'
'    oPCL.WriteTag "DOCUMENTCODE", Me.DocCode
'    oPCL.WriteTag "OURADDRESS", oPC.Configuration.DefaultCompany.StreetAddressDelimited
'    oPCL.WriteTag "DATE", Me.DocDateF
'    oPCL.WriteTag "DELTO", "TO: " & oPC.Configuration.Stores.FindStoreByID(DestID).Description
'
'    FillTableRows
'    oPCL.WriteTag "TOTAL1", TotalLessDiscExtF
'    oPCL.WriteTag "TOTAL2", TotalVATF
'    oPCL.WriteTag "TOTAL3", TotalPayableF
'    oPCL.WriteTag "TOTALDESCRIPTION", "subtotal" & "|" & "Less VAT of" & "|" & "Value"
'    oPCL.FinishRequest
'
'EXIT_Handler:
'    Exit Sub
'Err_Handler:
'    oPCL.FinishRequest
'    Select Case
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox error
'        GoTo EXIT_Handler
'        Resume
'    End Select
'End Sub
'Private Function FillTableRows()
'Dim i As Integer
'On Error GoTo err_Handler
'    For i = 1 To RLines.Count
'        With RLines(i)
'            oPCL.RLSend .CodeF, .Qty, .PriceF, .DiscountF, .PLessDiscExtF, .Title
'        End With
'    Next i
'EXIT_Handler:
'    Exit Function
'Err_Handler:
'    MsgBox error
'    GoTo EXIT_Handler
'   ' Resume
'End Function
Public Property Get ISForeignCurrency() As Boolean
    On Error GoTo errHandler
    ISForeignCurrency = (Not oPC.Configuration.DefaultCurrency.ID = Me.CurrencyID) And Me.CurrencyID > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ISForeignCurrency"
End Property

Public Function SetMemo(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Memo) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Memo Then
        mudtProps.Memo = val
        SetDirty True
    End If
    SetMemo = True
    Exit Function
End Function
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Memo"
End Property
Public Property Let Memo(val As String)
    mudtProps.Memo = Trim(val)
End Property


Public Property Get CaptureCurrency() As a_Currency
    Set CaptureCurrency = Supplier.DefaultCurrency
End Property
Public Property Let CurrencyID(ID As Long)
    mudtProps.CurrencyID = ID
End Property
Public Property Get CurrencyID() As Long
    CurrencyID = mudtProps.CurrencyID
End Property

Public Property Get CurrencyConversionAsText() As String

End Property
Public Function PrintReturn(bForeign As Boolean, pbIncZero As Boolean, Optional pQtyCopies As Integer) As Boolean
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim tmp As String
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl

    
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintReturn = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("RT_" & Me.DOCCode) = False Then
            PrintReturn = False
            Exit Function
        End If
        With oDC
        oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, IIf(pQtyCopies > 0, pQtyCopies, .QtyCopies)
        End With
        PrintReturn = True
    End If
    
    
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt RLines.Count
    oPCL.LineTotalCount 1
    oPCL.WriteTag "TOTAL", "Total value" & "|" & Me.TotalPayableF(bForeign)
    oPCL.WriteTag "QTYTOTAL", "Total qty" & "|" & Me.TotalQty
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "SUPPLIER", Supplier.NameAndCode(35)
    oPCL.WriteTag "STORECODE", oPC.Configuration.DefaultStore.code
    If Not Me.Supplier.BillTOAddress Is Nothing Then
        oPCL.WriteTag "PHONE", IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
        oPCL.WriteTag "FAX", IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
        oPCL.WriteTag "BILLTO", PackText(oPC.Configuration.DefaultStore.BillAddress)
    End If
    oPCL.WriteTag "DELTO", PackText(oPC.Configuration.DefaultStore.DelAddress)
    oPCL.WriteTag "DATE", Me.DOCDate
    oPCL.WriteTag "MEMO", PackText(Memo)
    oPCL.WriteTag "SENDER", oPC.Configuration.DefaultCompany.CompanyName
    oPCL.WriteTag "OURADDRESS", PackText(oPC.Configuration.DefaultCompany.StreetandPostalAddress)
    oPCL.WriteTag "MAINMSG", Me.ApprovalRef & "(" & Me.ApprovalTermDate & ")"  'PackText(oPC.Configuration.OrderText)
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    FillTableRows bForeign, pbIncZero
    
    oPCL.FinishRequest
    
EXIT_Handler:
'Err_Handler:
'    oPCL.FinishRequest
'    Select Case
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox error
'        GoTo EXIT_Handler
'        Resume
'    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.PrintReturn(bForeign)", bForeign
End Function
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
    On Error GoTo errHandler
'Dim strLabels As String
'Dim strValues As String
'    strLabels = ""
'    strValues = ""
'    If bForeign Then
'        If oPC.Configuration.IsVATRegion Then
'                strLabels = strLabels & Chr(13) & "Total"
'                strValues = strValues & Chr(13) & Me.TotalPayableF(bForeign)
'        Else
'                strLabels = strLabels & Chr(13) & "subtotal"
'                strValues = strValues & Chr(13) & Me.TotalExtF(bForeign)
'                strLabels = strLabels & Chr(13) & "Less discount"
'                strValues = strValues & Chr(13) & Me.TotalDiscountF((bForeign))
'                strLabels = strLabels & Chr(13) & "Total"
'                strValues = strValues & Chr(13) & Me.TotalPayableF(bForeign)
'        End If
'    Else
'        If oPC.Configuration.IsVATRegion Then
'                strLabels = strLabels & Chr(13) & "Total"
'                strValues = strValues & Chr(13) & Me.TotalExtF(bForeign)
'                strLabels = strLabels & Chr(13) & "Less discount"
'                strValues = strValues & Chr(13) & Me.TotalDiscountF((bForeign))
'                strLabels = strLabels & Chr(13) & "Includes VAT of "
'                strValues = strValues & Chr(13) & Me.TotalVATF(bForeign)
'        Else
'                strLabels = strLabels & Chr(13) & "subtotal"
'                strValues = strValues & Chr(13) & Me.TotalExtF(bForeign)
'                strLabels = strLabels & Chr(13) & "Less discount"
'                strValues = strValues & Chr(13) & Me.TotalDiscountF((bForeign))
'                strLabels = strLabels & Chr(13) & "Total"
'                strValues = strValues & Chr(13) & Me.TotalPayableF(bForeign)
'        End If
'    End If
'    pLabels = strLabels
'    pValues = strValues
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Sub FillTableRows(bForeign As Boolean, bIncZero As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim strQty As String


    For i = 1 To Me.RLines.Count
        With Me.RLines(i)
            If bIncZero Or .QtyReturned > 0 Or Me.Status <> stCOMPLETE Then
                If Me.StatusF = "AUTHORIZATION REQUESTED" Then
                    strQty = CStr(.QtyRequested)
                ElseIf Me.StatusF = "STOCK RETURNED" Then
                    strQty = CStr(.QtyReturned) & IIf(.QtyRejected > 0, " - " & CStr(.QtyRejected), "")
                End If
                oPCL.RLSend .CodeF, strQty, .Title, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountF, .Note, .SINVRef
            End If
        End With
    Next i
EXIT_Handler:
'Err_Handler:
'    MsgBox error
'    GoTo EXIT_Handler
   ' Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.FillTableRows(bForeign)", bForeign
End Sub

Public Function PrintReturn_Display(bForeign As Boolean) As String
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim ooRL As a_RL
Dim strR As String
Dim i As Integer

    
    strR = "Return: " & Me.DOCCode & vbCrLf
    
    For i = 1 To Me.RLines.Count
        With Me.RLines(i)
            strR = strR & .CodeF & "  " & .Title & "  " & .QtyReturned & "  " & Left(.Title, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.Discount = 0, "", .DiscountF) & "  " & .PLessDiscExtF(bForeign)
        End With
        strR = strR & vbCrLf
    Next i
    
    strR = strR & "Total   " & TotalPayableF(bForeign) & vbCrLf
    
    strR = strR & vbCrLf
   
    PrintReturn_Display = strR
EXIT_Handler:
'Err_Handler:
'    oPCL.FinishRequest
'    Select Case
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox error
'        GoTo EXIT_Handler
'        Resume
'    End Select
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.PrintReturn_Display(bForeign)", bForeign
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
    ErrorIn "a_R.VoidDocument"
End Sub
Public Property Get CanCancel() As Boolean
    CanCancel = True
End Property
Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.Reload"
End Sub

Public Function ExportToXML(bForeign As Boolean, pFilename As String, FilterQtygtZero As Boolean, _
        Optional DispatchMode As enTransmitType, Optional pQtyCopies As Integer, _
        Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, _
        Optional XMLFilePath As String, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim xslFilePath As String
Dim pdfFilePath As String
Dim FOPDebugPath As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim bReferenceExists As Boolean
Dim bDiscountExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim oCHRL As a_RL
Dim oDC As a_DocumentControl
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strTemplateName As String
Dim strExecutable As String
Dim lngNumberOfLines As Long
Dim oSM As z_StockManager
Dim oFSO As New FileSystemObject

    If DispatchMode = enMail Then
        If oPC.UsesOutlookForPOEmail Then
            strWorkingFolder = oPC.LocalFolder & "Emails\"
            strDispatchMethod = "Outlook"
        Else
            strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
            strDispatchMethod = "Emailing"
        End If
    ElseIf DispatchMode = enPrint Then
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Then
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
'        MsgBox "Change here - call support"
'        strWorkingFolder = oPC.LocalFolder & "PDF\"
        
        strDispatchMethod = ""
    Else
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If
    
                        p 1
    If bForeign = True Then
        CalculateTotal_Foreign
    End If
                        p 2
    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
    bQtyColumnNeeded = False
    For Each oCHRL In RLines
        If oCHRL.SINVRef > " " Then
            bReferenceExists = True
        End If
        If oCHRL.Discount > 0 Then
            bDiscountExists = True
        End If
        bQtyColumnNeeded = True
    Next
                        p 3
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
        ExportToXML = False
        Exit Function
    End If
                        p 4
    If pQtyCopies < 1 Then pQtyCopies = 1
    
    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "SR_DOC"
            .chCreate "MessageType"
                .elText = "RETURN"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "DestinationAddress"
                If oPC.TestMode Then
                    .elText = oPC.EmailAddressForTesting
                    pDestinationEmailAddress = oPC.EmailAddressForTesting
                Else
                    If Me.Supplier Is Nothing Then
                        .elText = ""
                        pDestinationEmailAddress = ""
                    Else
                        If Me.Supplier.OrderToAddress Is Nothing Then
                            .elText = ""
                            pDestinationEmailAddress = ""
                        Else
                            .elText = Me.Supplier.OrderToAddress.EMail
                            pDestinationEmailAddress = Me.Supplier.OrderToAddress.EMail
                        End If
                    End If
                End If
                        p 23
            .elCreateSibling "TemplateName"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\SR_DOC.XSLT") Then
                    .elText = "SR_DOC.XSLT"
                Else
                    If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\SR_DOC.XSL") Then
                        .elText = "SR_DOC.XSL"
                    Else
                        .elText = "SR_DOC_FO.XSL"
                    End If
                End If
                strTemplateName = .elText
            .elCreateSibling "SendersEmail"
                If oPC.EmailFrom > "" Then
                    .elText = oPC.EmailFrom
                Else
                    .elText = ""
                End If
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusF
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailPOMsg
            .elCreateSibling "DispatchMethod"
                .elText = Me.Supplier.DispatchModes.ItemByOrdinalIndex(Me.Supplier.DispatchModes.FindIndexByKey(Me.DispatchModeID))
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
                'THE FOLLOWING IS VITAL
             .elText = Replace(.elText, "\", "/")
             
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
            .elCreateSibling "DocDate", True
                .elText = Me.DOCDate
            .elCreateSibling "ApprovalRef"
                .elText = Me.ApprovalRef & "(" & Me.ApprovalTermDate & ")"
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "SupplierName", True
                .elText = Supplier.NameAndCode(35)
            .elCreateSibling "SupplierWithAddress", True
                If Supplier.OrderToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
            .elCreateSibling "SupplierPhone", True
                If Supplier.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "SupplierFax", True
                If Supplier.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
                End If
            
            
            
            .elCreateSibling "ACNO"
                .elText = IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
            .elCreateSibling "BillTo", True
                If Supplier.BillTOAddress.ID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Supplier.BillTOAddress.ID) Is Nothing Then
                        .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Supplier.BillTOAddress.ID).BillAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        .elText = Replace(oPC.Configuration.DefaultStore.BillAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
            .elCreateSibling "DelTo", True
                If Supplier.DelToAddress.ID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Supplier.DelToAddress.ID) Is Nothing Then
                        .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Supplier.DelToAddress.ID).DelAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        .elText = Replace(oPC.Configuration.DefaultStore.DelAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
                            p 8
                            
                            
                            
'    For i = 1 To Me.RLines.Count
'        With Me.RLines(i)
'            If bIncZero Or .QtyReturned > 0 Or Me.STATUS <> stCOMPLETE Then
'                If Me.StatusF = "AUTHORIZATION REQUESTED" Then
'                    strQty = CStr(.QtyRequested)
'                ElseIf Me.StatusF = "STOCK RETURNED" Then
'                    strQty = CStr(.QtyReturned) & IIf(.QtyRejected > 0, " - " & CStr(.QtyRejected), "")
'                End If
'                oPCL.RLSend .CodeF, strQty, .Title, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountF, .Note, .SINVRef
'            End If
'        End With
'    Next i
                            
                            
            lngNumberOfLines = 0
            For i = 1 To RLines.Count
                If (Me.Status = stCOMPLETE And RLines.Item(i).QtyReturned > 0) Or _
                    (Me.Status = stISSUED And RLines.Item(i).QtyRequested > 0) Or FilterQtygtZero = False Or Me.Status = stInProcess Then
                    lngNumberOfLines = lngNumberOfLines + 1
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                     .elText = RLines.Item(i).CodeF
                    .elCreateSibling "Title", True
                        .elText = RLines.Item(i).Title '& IIf(RLines(i).ServiceItem, " (" & RLines(i).Qty & " item" & IIf(RLines(i).Qty = 1, ")", "s)"), "")
                     .elCreateSibling "Qty", True
                    If Status = stCOMPLETE Then
                        .elText = RLines.Item(i).QtyReturned
                    Else
                        .elText = RLines.Item(i).QtyRequested
                    End If
                        
                    .elCreateSibling "Price", True
                        .elText = RLines.Item(i).PriceF(bForeign)
                    .elCreateSibling "Reference", True
                        .elText = RLines.Item(i).SINVRef
                    .elCreateSibling "Extension", True
                        .elText = RLines.Item(i).PLessDiscExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = RLines.Item(i).Note
                    .elCreateSibling "DiscountRate", True
                        .elText = RLines(i).DiscountF
                    .navUP
                End If
            Next
                                p 10
            .elCreateSibling "TotalNumberOfLines", True
                .elText = CStr(lngNumberOfLines)
            .elCreateSibling "QtyItems", True
                .elText = Me.TotalQty
                                p 11

            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "StaffMember", True
                .elText = IIf(Me.StaffName > "", Me.StaffName, "")
            .elCreateSibling "Memo", True
                .elText = "" 'oPC.Configuration.returntext
                
            .elCreateSibling "TotalIncVAT", True
                .elText = Me.TotalPayableF(bForeign)
    End With
                                p 12

'FINALLY PRODUCE THE .XML FILE
    XMLFilePath = strWorkingFolder & "RT_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(XMLFilePath) Then
            fs.DeleteFile XMLFilePath
        End If
        .docWriteToFile (XMLFilePath), False, "UNICODE", "" 'strHead
    End With
    If Not fs.FileExists(XMLFilePath) Then
        MsgBox "Cannot write XML file.", vbOKOnly, "Can't do this"
        Exit Function
    End If
    
                p 13

    If DispatchMode = enMail Then
        If oPC.EmailPOShowHTML = True Then
            'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_HTML.XSLT"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
            If fs.FileExists(strWorkingFolder & "SR_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "SR_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "R_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage
            oTF.CloseTextFile
        End If
    End If
    
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    Dim tmp As String
    If (Me.Supplier.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
            xslFilePath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName
        pdfFilePath = strWorkingFolder & "PO_" & Me.DOCCode & ".PDF"
        FOPDebugPath = strWorkingFolder & "FOPDebug.txt"
        If UCase(Right(strTemplateName, 4)) = "XSLT" Then
          tmp = oPC.LocalFolder & "Executables\FOP-1.1\"
        Else
          tmp = oPC.LocalFolder & "Executables\FOP\"
        End If
        strCommand = tmp & "FOP.BAT  -xml " & XMLFilePath & " -xsl " & xslFilePath & " -pdf " & pdfFilePath & " 2>" & FOPDebugPath
                LogSaveToFile "strCommand = " & strCommand
           'Allow 2 minutes to produce the PDF file
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000

'        Set objXSL = Nothing
'        Set objXSL = New MSXML2.DOMDocument60
'        objXSL.async = False
'        objXSL.ValidateOnParse = False
'        objXSL.resolveExternals = False
'        strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "SR_DOC_FO.XSL"
''        MsgBox "Change here too"
''        strPath = oPC.LocalFolder & "Templates\" & strTemplateName & "SR_DOC_FO.XSL"
'        Set fs = New FileSystemObject
'        If fs.FileExists(strPath) Then
'            objXSL.Load strPath
'        End If
'                        p 6
'        Set opXMLDOC = New MSXML2.DOMDocument60
'        opXMLDOC.async = False
'        opXMLDOC.ValidateOnParse = False
'        opXMLDOC.resolveExternals = False
'        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
'
'        strFOFile = strWorkingFolder & "SR_" & Me.DOCCode & ".FO"
'        strPDFFile = strWorkingFolder & "SR_" & Me.DOCCode & ".PDF"
'                        p 7
'        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
'
''Stage 2 Convert the .FO file to .PDF and clean up
'        ChDir "\PBKS\Executables\FOP\"
'       ' strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
'        'ShellandWait strCommand, 100
'        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
'        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
'
'        If fs.FileExists(strFOFile) Then
'            fs.DeleteFile strFOFile
'        End If
    End If
    
'Send to Outlook
   ' CreateOutlookMail strXML
    
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    pFilename = pdfFilePath
    
    If DispatchMode <> enMail And DispatchMode <> enView And DispatchMode <> enXML Then
        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    End If
    Set oSM = Nothing
    Set oFSO = Nothing
                        p 9
    ExportToXML = True

    '
'    If DispatchMode = enView Then
'        strExecutable = GetPDFExecutable(strPDFFile)
'                If strExecutable = "" Then
'                    MsgBox "There is no application set on this computer to open the file: " & strPDFFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
'                Else
'                    Shell strExecutable & " " & strPDFFile, vbNormalFocus
'                End If
'    End If
    If (Not fs.FileExists(pdfFilePath)) And (DispatchMode = enView Or DispatchMode = enMail) Then
            MsgBox "The PDF file has not been produced: " & pdfFilePath & ". Please contact support.", vbOKOnly, "Can't do this"
    Else
        If DispatchMode = enView Or DispatchMode = enPrint Then
            OpenFileWithApplication pdfFilePath, enPDF, QuickPrint
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
End Function

Public Function ExportToSpreadsheet(bForeign As Boolean, _
        pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim s As String
Dim s2 As String
Dim lngNumberOfLines As Long

    ExportToSpreadsheet = False
    If bForeign = True Then
        CalculateTotal_Foreign
    End If
    pFilename = oPC.LocalFolder & "Temp\" & DOCCode & "_" & Format(Now(), "yyyymmddHHnn") & ".xls"
    
    oTF.OpenTextFile pFilename
    oTF.WriteToTextFile "Status" & vbTab & StatusF
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCode
    oTF.WriteToTextFile "DocDate" & vbTab & DOCDate
    oTF.WriteToTextFile "ApprovalRef" & vbTab & ApprovalRef & "(" & Me.ApprovalTermDate & ")"
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.DefaultCompany.CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "SupplierName" & vbTab & Supplier.NameAndCode(70)
    oTF.WriteToTextFile_NoLineTerminator "SupplierWithAddress" & vbTab
                If Not Supplier.OrderToAddress Is Nothing Then
                    oTF.WriteToTextFile Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
    
    oTF.WriteToTextFile_NoLineTerminator "SupplierPhone" & vbTab
                    oTF.WriteToTextFile IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
    oTF.WriteToTextFile_NoLineTerminator "SupplierFax" & vbTab
                    oTF.WriteToTextFile IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
    oTF.WriteToTextFile "ACNO" & vbTab & IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
            
        If Supplier.BillTOAddress.ID > 0 Then
            If Not oPC.Configuration.Stores.FindStoreByID(Supplier.BillTOAddress.ID) Is Nothing Then
                s = Replace(oPC.Configuration.Stores.FindStoreByID(Supplier.BillTOAddress.ID).BillAddress, Chr(13) & Chr(10), Chr(10))
            Else
                s = Replace(oPC.Configuration.DefaultStore.BillAddress, Chr(13) & Chr(10), Chr(10))
            End If
        End If
    oTF.WriteToTextFile "BillTo" & vbTab & s
            
        If Supplier.DelToAddress.ID > 0 Then
            If Not oPC.Configuration.Stores.FindStoreByID(Supplier.DelToAddress.ID) Is Nothing Then
                s = Replace(oPC.Configuration.Stores.FindStoreByID(Supplier.DelToAddress.ID).DelAddress, Chr(13) & Chr(10), Chr(10))
            Else
                s = Replace(oPC.Configuration.DefaultStore.DelAddress, Chr(13) & Chr(10), Chr(10))
            End If
        End If
    oTF.WriteToTextFile "DelTo" & vbTab & s
                            
    s = "SKU" & vbTab & "Title" & vbTab & "Qty" & vbTab & "Price" & "Reference" & vbTab & "Extension" & vbTab & "DiscountRate" & vbTab & "Note"
    oTF.WriteToTextFile s
        lngNumberOfLines = 0
        For i = 1 To RLines.Count
                lngNumberOfLines = lngNumberOfLines + 1
                s = RLines.Item(i).CodeF & vbTab _
                & RLines.Item(i).Title & vbTab
                
                If Status = stCOMPLETE Then
                    s2 = RLines.Item(i).QtyReturned
                Else
                    s2 = RLines.Item(i).QtyRequested
                End If
                s = s & s2
                s = s & vbTab & RLines.Item(i).PriceF(bForeign)
                s = s & vbTab & RLines.Item(i).SINVRef
                s = s & vbTab & RLines.Item(i).PLessDiscExtF(bForeign)
                s = s & vbTab & RLines(i).DiscountF
                s = s & vbTab & RLines.Item(i).Note
                oTF.WriteToTextFile s
            Next
            
    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(lngNumberOfLines)
    oTF.WriteToTextFile "QtyItems" & vbTab & TotalQty
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.WriteToTextFile "StaffMember" & vbTab & IIf(Me.StaffName > "", Me.StaffName, "")
    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalPayableF(True)
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
End Function




Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "a_R.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub
Public Property Get StatusForPrinting() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusForPrinting = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusForPrinting = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusForPrinting = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.StatusForPrinting"
End Property

Public Property Get DispatchModeID() As Long
    DispatchModeID = FNS(mudtProps.DispatchMethodID)
End Property
Public Function SetDispatchModeID(val As Long) As Boolean
Dim bOK As Boolean
        mudtProps.DispatchMethodID = val
End Function

Public Property Get Log() As String
    On Error GoTo errHandler
    Log = Trim(mudtProps.Log)
    Exit Property
errHandler:
    ErrorIn "a_R.Log"
End Property

Public Property Let Log(val As String)
    On Error GoTo errHandler
    mudtProps.Log = Trim(val)
    Exit Property
errHandler:
    ErrorIn "a_R.Log(val)", val
End Property



