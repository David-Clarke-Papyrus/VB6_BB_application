VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngTotal1 As Long
Dim lngTotal2 As Long
Dim lngTotal3 As Long
Dim lngTotal4 As Long
Dim lngTotal5 As Long
Dim lngTotalDeposit As Long
Dim lngTotalVAT As Long
Dim lngTotalInvoiceDiscount As Long
Dim lngTotalFinal As Long
Dim lngTotalDiscount As Long
Dim lngTotalExtras As Long
Dim lngTotalExtensionInclDepositIncDiscount As Long
Dim oPCL As z_PrintClient
Dim mOriginalStatus As enStatus
Private xMLDoc As ujXML

Event Valid(pMsg As String)
Private mudtProps As APPProps
Private mcolStack As Collection
Private mColApproLines As ch_APPL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private colClassors As Collection
Private oApproAddress As a_Address
Private oBillingCompany As a_Company
Const CDOCCODE = "AP"
'Type tTotal
'    Description As String
'    val As Long
'    AmtFormatted As String
'    TotalType As String
'    RunningTotal As Long
'    Sign As String
'End Type
Dim arDocTotals() As tTotal
Dim arDocTotalsF() As tTotal
Dim RunningTotal As Long
Dim jTotal As Long

Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set mColApproLines = New ch_APPL
    mColApproLines.component Me
    Set mCustomer = New a_Customer
    Set oApproAddress = New a_Address
    mobjValid.BreakRule "TP", True
    Set mcolStack = New Collection
    mOriginalStatus = stInProcess
    LoadClassorsCollection
    mudtProps.COMPID = oPC.Configuration.DefaultCompany.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set mobjValid = Nothing
    Set mColApproLines = Nothing
    Set mCustomer = Nothing
    Set oApproAddress = Nothing
    Set mcolStack = Nothing
    Set colClassors = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(ors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(ors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.mobjValid_BrokenRule(ors)", ors
End Sub
Private Sub mobjValid_RuleUnbroken(ors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(ors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.mobjValid_RuleUnbroken(ors)", ors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.mobjValid_Status(pMsg)", pMsg
End Sub
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As APPData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.GetState"
End Function
Private Sub SetState(Buffer As String)
    On Error GoTo errHandler
  Dim udtData As APPData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.SetState(buffer)", Buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "ALS", mColApproLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.GetSuperState"
End Function
Public Sub SetSuperState(Buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = Buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mColApproLines.SetSuperState .ReadProperty("ALS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.SetSuperState(buffer)", Buffer
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mColApproLines.BeginEdit
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mColApproLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.CancelEdit"
End Sub

Public Sub ApplyEdit(bMadeChange As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_APP_P
  If mcolStack.Count = 0 Then Err.Raise 445
    If DOCCode = "" Then
        DOCCode = GetAPPNumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now

  mColApproLines.ApplyEdit
  Set objPersist = New a_APP_P
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject mudtProps.TRID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    mcolStack.Remove mcolStack.Count
    SetSuperState objPersist.Save(GetSuperState, bMadeChange)
    mudtProps.IsNew = False
  Else
    mcolStack.Remove mcolStack.Count
  End If
  Set objPersist = Nothing
  mudtProps.IsDirty = False
  Dim lngAPPID As Long
  lngAPPID = Me.TRID
  'THis is necessary because we consolidate rows on saving
   ' Reload
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.ApplyEdit"
End Sub
Public Function GetAPPNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strWork As String

    Set oCode = New z_Code
    
    strWork = oCode.GetNextCode(enAppro, strTmp)
    strDefaultCo = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyCode
    
    GetAPPNumber = strDefaultCo & "A" & strWork
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.GetAPPNumber(strTmp)", strTmp
End Function

Public Sub Load(AppID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_APP_P
  
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    Set objPersist = New a_APP_P
    SetSuperState objPersist.Fetch(AppID)
    Set objPersist = Nothing
    
    SetCustomer Me.TPID
    If Me.APPROTOID > 0 Then
        Set oApproAddress = Customer.Addresses.FindByID(Me.APPROTOID)
        'The address might not be found if it has been removed or the customer has been merged with another
        If oApproAddress Is Nothing Then
            Set oApproAddress = Customer.ApproAddress
            If Not oApproAddress Is Nothing Then
                Me.APPROTOID = Customer.ApproAddress.ID
            End If
        End If
    End If
    mOriginalStatus = Me.STATUS
    mudtProps.IsNew = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Load(AppID,ReadOnly)", Array(AppID, ReadOnly)
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mColApproLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IsValid"
End Property
Public Function CalculateTotal()
    On Error GoTo errHandler
Dim oAPPL As a_APPL
    mobjValid.RuleBroken "OVERFLOW", False
    lngTotal1 = 0
    lngTotal2 = 0
    lngTotal3 = 0
    lngTotal4 = 0
    lngTotal5 = 0
    For Each oAPPL In mColApproLines
  '      oAPPL.RecalculateLine

        lngTotal1 = lngTotal1 + oAPPL.ExtensionSimple
        lngTotal2 = lngTotal2 + oAPPL.ExtensionNet
        lngTotal3 = lngTotal3 + oAPPL.ExtensionNetExVAT
        lngTotal4 = lngTotal4 + oAPPL.ExtensionVATOnly
        lngTotal5 = lngTotal5 + oAPPL.Qty
    Next
    mudtProps.TotalGross = lngTotal1
    mudtProps.TotalNet = lngTotal2
    mudtProps.TotalNetExVAT = lngTotal3
    mudtProps.TotalVAT = lngTotal4
    mudtProps.TotalQty = lngTotal5
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.CalculateTotal"
End Function
Public Property Get TotalGross() As Long
    On Error GoTo errHandler
    TotalGross = mudtProps.TotalGross
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalGross"
End Property
Public Property Get TotalGrossF() As String
    On Error GoTo errHandler
    TotalGrossF = Format(mudtProps.TotalGross / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalGrossF"
End Property
Public Property Get TotalNet() As Long
    On Error GoTo errHandler
    TotalNet = mudtProps.TotalNet
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalNet"
End Property
Public Property Get TotalNetF() As String
    On Error GoTo errHandler
    TotalNetF = Format(mudtProps.TotalNet / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalNetF"
End Property
Public Property Get TotalNetExVAT() As Long
    On Error GoTo errHandler
    TotalNetExVAT = mudtProps.TotalNetExVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalNetExVAT"
End Property
Public Property Get TotalNetExVATF() As String
    On Error GoTo errHandler
    TotalNetExVATF = Format(mudtProps.TotalNetExVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalNetExVATF"
End Property
Public Property Get TotalVAT() As Long
    On Error GoTo errHandler
    TotalVAT = mudtProps.TotalVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalVAT"
End Property
Public Property Get TotalVATF() As String
    On Error GoTo errHandler
    TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalVATF"
End Property
Public Property Get TotalQtyItems() As Long
    On Error GoTo errHandler
    TotalQtyItems = mudtProps.TotalQty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TotalQtyItems"
End Property

Public Property Get TPName() As String
    On Error GoTo errHandler
    TPName = FNS(mudtProps.TPName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TPName"
End Property

Public Property Let TPName(Value As String)
    On Error GoTo errHandler
    mudtProps.TPName = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TPName(Value)", Value
End Property

Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.STATUS = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.STATUS = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.STATUS = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.STATUS = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.STATUS = stPROFORMA Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.StatusF"
End Property
Public Property Get STATUS() As enStatus
    On Error GoTo errHandler
    STATUS = mudtProps.STATUS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Status"
End Property

Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Memo"
End Property
Public Property Let Memo(val As String)
    On Error Resume Next
    mudtProps.Memo = Trim(val)
    SetDirty True
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
    
Public Property Let IssDate(val As Date)
    On Error GoTo errHandler
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IssDate(val)", val
End Property
Public Property Get IssDateF() As String
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IssDateF"
End Property

Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IssDate"
End Property
    
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.DocDateF"
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.DocDate"
End Property

Public Property Let DOCDate(Value As Date)
    On Error GoTo errHandler
    If IsDate(Value) Then
        mudtProps.DOCDate = CDate(Value)
        mudtProps.IsDirty = True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.DocDate(Value)", Value
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.DocCode"
End Property

Public Property Let DOCCode(Value As String)
    On Error GoTo errHandler
    mudtProps.DOCCode = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.DocCode(Value)", Value
End Property

Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TPID"
End Property

Public Property Let TPID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TPID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TPID(Value)", Value
End Property

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TRID"
End Property

Public Property Let TRID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TRID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TRID(Value)", Value
End Property
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.StaffID"
End Property

Public Property Let StaffID(Value As Long)
    On Error GoTo errHandler
    mudtProps.StaffID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.StaffID(Value)", Value
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
    ErrorIn "a_APP.StaffName"
End Property

Public Property Get Shortname() As String
    On Error GoTo errHandler
    If mudtProps.StaffID < 1 Then
        Shortname = ""
        Exit Property
    End If
    Shortname = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Shortname"
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
    ErrorIn "a_APP.StaffNameB"
End Property

Public Property Get APPROTOID() As Long
    On Error GoTo errHandler
    APPROTOID = mudtProps.APPROTOID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.APPROTOID"
End Property

Public Property Let APPROTOID(Value As Long)
    On Error GoTo errHandler
    mudtProps.APPROTOID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.APPROTOID(Value)", Value
End Property

Public Function GetTotalValue() As Long
    On Error GoTo errHandler
Dim oAL As a_APPL
Dim curTotalValue As Long
    curTotalValue = 0
    For Each oAL In Me.ApproLines
        curTotalValue = curTotalValue + (((oAL.Qty - oAL.QtyReturned) * oAL.Price) * ((100 - oAL.Discount) / 100))
    Next
    GetTotalValue = curTotalValue
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.GetTotalValue"
End Function
Public Function GetTotalValueF() As String
    On Error GoTo errHandler
    GetTotalValueF = Format(GetTotalValue / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.GetTotalValueF"
End Function
Public Property Get ApproLines() As ch_APPL
  Set ApproLines = mColApproLines
End Property

Public Property Get DisplayAppro() As d_OSAPPRO
    On Error GoTo errHandler
    Set DisplayAppro = New d_OSAPPRO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.DisplayAppro"
End Property
Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo errHandler
    mudtProps.TPACCNum = val
  mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.TPACCNum(val)", val
End Property
Public Property Get COMPID() As Long
    On Error GoTo errHandler
    If mudtProps.COMPID > 0 Then
        COMPID = mudtProps.COMPID
    Else
        COMPID = oPC.Configuration.DefaultCOMPID
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.COMPID"
End Property
Public Property Let COMPID(val As Long)
    On Error GoTo errHandler
Dim oComp As a_Company
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COMPID = val
    SetDirty True
    For Each oComp In oPC.Configuration.Companies
        If oComp.ID = mudtProps.COMPID Then
            Set oBillingCompany = oComp
        End If
    Next
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.COMPID(val)", val
End Property

Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.VATRate"
End Property

Public Property Let VATRate(Value As Double)
    On Error GoTo errHandler
    mudtProps.VATRate = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.VATRate(Value)", Value
End Property

Public Property Get VATable() As Boolean
    On Error GoTo errHandler
  VATable = Me.Customer.VATable
 ' VATable = mudtProps.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATable(val)", val
End Property
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ShowVAT"
End Property
Public Property Let ShowVAT(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowVAT = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ShowVAT(val)", val
End Property

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.SetDirty(pVal)", pVal
End Sub
Public Function Post() As String
    On Error GoTo errHandler
Dim strResult As String
Dim oSM As z_StockManager
Dim bMadeChange As Boolean
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    Me.ApplyEdit bMadeChange
    
    Set oSM = New z_StockManager
    
    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
            Select Case STATUS
            Case stVOID
       ''         SetStatus stVOID
            Case stInProcess
                Post = "In Process"
            Case stCOMPLETE
                oSM.IssueAPP Me
            Case stISSUED
                oSM.IssueAPP Me
              '  SetStatus stCOMPLETE
            Case stPROFORMA
                oSM.IssueAPP Me
        '        SetStatus stPROFORMA
            End Select
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    Case stPROFORMA      'Issued PROFORMA
            Select Case STATUS
            Case stVOID
                oSM.CancelAPP Me
        '        SetStatus stVOID
            Case stInProcess
                Post = "OR: was proforma, can't now be inprocess."
            Case stCOMPLETE
                oSM.IssueAPP Me
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    Set oSM = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    If bMadeChange Then Reload  'In case lines have been consolidated

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Post", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function

Public Function PrintAPP(Optional pQtyCopies As Integer) As Boolean
    On Error GoTo errHandler
Dim iTotalLinesCOunt As Integer
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl

    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintAPP = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("AP_" & Me.DOCCode) = False Then
            PrintAPP = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, IIf(pQtyCopies > 0, pQtyCopies, .QtyCopies)
        End With
        PrintAPP = True
    End If
    
    
    
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt ApproLines.Count
    iTotalLinesCOunt = 0
    oPCL.LineTotalCount 2
 '   oPCL.LineTotalCount iTotalLinesCOunt 'This must be written before the TOTAL tags for the Printserver requirements
    oPCL.WriteTag "TOTAL", "Total" & "|" & Me.TotalNetF
    oPCL.WriteTag "TOTAL", "Includes V.A.T. of" & "|" & Me.TotalVATF
    oPCL.WriteTag "TOTAL", ""
    oPCL.WriteTag "TOTAL", "Total qty items|" & Me.TotalQtyItems
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "LOGO", oPC.SharedFolderRoot & "\LOGOS\" & oPC.Configuration.DefaultCompany.LogoFilePath
'    oPCL.WriteTag "OURADDRESS", oPC.Configuration.DefaultCompany.StreetAddressDelimited
    oPCL.WriteTag "SENDER", oPC.Configuration.Companies.FindCompanyByID(COMPID).CompanyName
    oPCL.WriteTag "OURADDRESS", PackText(oPC.Configuration.Companies.FindCompanyByID(COMPID).StreetandPostalAddress)
    oPCL.WriteTag "COMPANYNAME", PackText(oPC.Configuration.DefaultStore.BillAddress)
    If Not ApproToAddress Is Nothing Then
        oPCL.WriteTag "PHONE", ApproToAddress.PhoneandFax
    End If
  '  oPCL.WriteTag "FAX", IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
    oPCL.WriteTag "COREGNUMBER", oPC.Configuration.DefaultCompany.CoRegistrationNumber
    oPCL.WriteTag "VATNUMBER", oPC.Configuration.DefaultCompany.VatNumber
    oPCL.WriteTag "CUSTOMER", Me.Customer.NameAndCode(25)
    oPCL.WriteTag "DATE", Me.IssDateF
    oPCL.WriteTag "MEMO", PackText(Memo)
    oPCL.WriteTag "STORECODE", oPC.Configuration.DefaultStore.code
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    If Not ApproToAddress Is Nothing Then
        oPCL.WriteTag "BILLTO", Me.ApproToAddress.AddressDelimited
    End If

    FillTableRows
    oPCL.FinishRequest

EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.PrintAPP"
End Function
Private Function FillTableRows()
    On Error GoTo errHandler
Dim i As Integer
Dim oLine As a_APPL

    For i = 1 To ApproLines.Count
        Set oLine = Me.ApproLines(i)
        With oLine
            oPCL.APPLSend .CodeF, .Qty, .Description, .Ref, .PriceF, .ExtensionNetF, .ExtensionNetExVATF, .DiscountF, PackText(.Note) '"InvDate", oInv.TransDateFormatted
        End With
    Next i
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.FillTableRows"
End Function



Public Property Get Customer() As a_Customer
    Set Customer = mCustomer
End Property
Public Sub SetApproToAddress(val As a_Address)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    Set oApproAddress = val
    mudtProps.APPROTOID = oApproAddress.ID
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.SetApproTAddress(val)", val
End Sub
Property Get ApproToAddress() As a_Address
    Set ApproToAddress = oApproAddress
End Property
Public Function SetCustomer(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    mudtProps.TPID = pTPID
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
        Set oApproAddress = Me.Customer.ApproAddress
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.SetCustomer(pTPID)", pTPID
End Function

Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.STATUS = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.SetStatus(val)", val
End Sub
Public Function GetSTatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.GetSTatus"
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
    ErrorIn "a_APP.TranslateErrors(pRawors)", pRawors
End Function
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.IsEditing"
End Property
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.LoadClassorsCollection"
End Sub

Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.Reload"
End Sub
Public Property Get CanCancel() As Boolean
    On Error GoTo errHandler
Dim bCanCancel As Boolean
Dim oAPPL As a_APPL
    bCanCancel = True
    For Each oAPPL In ApproLines
        If oAPPL.QtyReturned > 0 Then
            bCanCancel = False
            Exit For
        End If
    Next
    CanCancel = bCanCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.CanCancel"
End Property

Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
Dim bMadeChanges As Boolean

    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit bMadeChanges
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.VoidDocument"
End Function

Public Function PrintSlips() As Boolean
    On Error GoTo errHandler
Dim iTotalLinesCOunt As Integer
Dim oSM As New z_StockManager
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

    Set oPCL = New z_PrintClient
    If oPCL.StartRequestPrinting("APS_" & Me.DOCCode) = False Then
        PrintSlips = False
        Exit Function
    End If
    With oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
        oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
    End With
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
 '   Set rs = oSM.GetApproSlips(Me.TRID)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM vAPPROSLIPS WHERE TR_ID = " & TRID & " ORDER BY AUTHOR", oPC.COShort, adOpenForwardOnly, adLockReadOnly

    oPCL.LineCOunt rs.RecordCount

    Do While Not rs.eof
        oPCL.APPSLIPSend FNS(rs!TP_Name), rs!CD, FNS(rs!code), FNS(rs!Title), FNS(rs!TR_CODE), _
            rs!dte, FNS(rs!ADD_Phone), FNS(rs!ADD_Fax), FNS(rs!TP_Initials), FNS(rs!TP_Title), FNS(rs!ADD_BusPhone), _
            rs!TP_Vatable, PackText(FNS(rs!ADD_L1)), FNS(rs!Author), FNS(rs!TR_NOTE), FNS(rs!TP_ACno), rs!APPL_Qty, PBKSPercentF(rs!Discount), _
            Format(CDbl(rs!APPL_Price) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
            rs!TR_ID, FNS(rs!Publisher), FNS(rs!COMP_Name), PackText(FNS(rs!COMP_StreetAdd)), _
            Format((CDbl(rs!APPL_Price)) * ((100# - rs!Discount) / 100) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
            oPC.Configuration.DefaultCompany.VatNumber <> ""
        rs.MoveNext
        '*(100 / (100 + IIf(oPC.Configuration.DefaultCompany.VatNumber = "", FNDBL(rs!APPL_VATRATE), 0))))
        ' * (100 / (100 + IIf(C, FNDBL(rs!APPL_VATRATE), 0))) * ((100 - rs!Discount) / 100)
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    oPCL.FinishRequest

EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP.PrintSlips"
End Function

Public Sub PasteLine(PID As String, Qty As Long, Price As Long, DiscountRate As Double, VATRate As Double, Ref As String)
Dim oSM As New z_StockManager

    oSM.PasteApproLine PID, Qty, Price, DiscountRate, VATRate, mudtProps.TRID, Ref
    
End Sub

Public Property Get SendersEmail()
    SendersEmail = FNS(mudtProps.StaffEmail)
End Property
Private Property Get StatusForPrinting() As String
    On Error GoTo errHandler
        
    If mudtProps.STATUS = stVOID Then
        StatusForPrinting = "VOID"
    ElseIf mudtProps.STATUS = stInProcess Then
        StatusForPrinting = "IN PROCESS"
    ElseIf mudtProps.STATUS = stISSUED Then
        StatusForPrinting = ""
    ElseIf mudtProps.STATUS = stCOMPLETE Then
        StatusForPrinting = ""
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusForPrinting = "CANCELLED"
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.StatusForPrinting"
End Property



'================================================================
Public Function ExportToXML(DispatchMode As enTransmitType, _
            Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, _
            Optional pQtyCopies As Integer, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
Dim bForeign As Boolean
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strXML As String
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
'Dim oCHKIL As a_QUL
Dim oDC As a_DocumentControl
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strTemplateName As String
Dim strExecutable As String

    If DispatchMode = enMail Then
        If oPC.UsesOutlookForINVEmail Then
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
        strDispatchMethod = ""
    Else
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If
    
                        p 1
    bForeign = False
                        p 2
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
        .docInit "APP_DOC"
            .chCreate "MessageType"
                .elText = "APPRO"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "DestinationAddress"
                If oPC.TestMode Then
                    .elText = oPC.EmailFrom
                Else
                    If Customer.BillTOAddress Is Nothing Then
                        .elText = ""
                    Else
                        .elText = Customer.BillTOAddress.EMail
                    End If
                End If
            .elCreateSibling "TemplateName"
                If Me.Customer.ApproTemplateName > "" Then
                    .elText = Me.Customer.ApproTemplateName
                Else
                    .elText = "APP_DOC"
                End If
            strTemplateName = .elText
            
                        p 5
            .elCreateSibling "RecipientsEmail"
                If Not Me.Customer.BillTOAddress Is Nothing Then
                    .elText = Me.Customer.BillTOAddress.EMail
                    pDestinationEmailAddress = Me.Customer.BillTOAddress.EMail
                End If
            .elCreateSibling "SendersEmail"
                .elText = Me.SendersEmail
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailAPPMsg
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
           .elCreateSibling "StoreCode", True
                .elText = oPC.Configuration.DefaultStore.code
           .elCreateSibling "DocDate", True
                .elText = Me.DocDateF
            .elCreateSibling "OurCompany", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
                                p 6
            .elCreateSibling "OurCompanyAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerName", True
                .elText = Customer.Name
            
            .elCreateSibling "OurAcnowithCustomer", True
                .elText = Customer.OurACnoWithClient
            
            
            .elCreateSibling "CustomerVATNumber", True
            If Customer.VatNumber > "" Then
                .elText = Customer.VatNumber
            Else
                .elText = ""
            End If
            .elCreateSibling "ACNO", True
                .elText = Customer.AcNo
            .elCreateSibling "BankingDetails", True
                .elText = Replace(oPC.Configuration.DefaultCompany.BankDetails, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerWithAddress", True
                If Me.ApproToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Me.ApproToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
                                p 7
            .elCreateSibling "CustomerPhone", True
                If Me.ApproToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.ApproToAddress.Phone > "", "Phone: " & Me.ApproToAddress.Phone, "")
                End If
            .elCreateSibling "CustomerFax", True
                .elText = IIf(Me.ApproToAddress.Fax > "", "Fax: " & Me.ApproToAddress.Fax, "")
            .elCreateSibling "BillTo", True
                If Me.ApproToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(ApproToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.ApproToAddress.Phone > "", "Phone: " & Me.ApproToAddress.Phone, "") & Chr(10) & IIf(Me.ApproToAddress.Fax > "", "Fax: " & Me.ApproToAddress.Fax, "")
                End If
                            p 8

            For i = 1 To Me.ApproLines.Count
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    If oPC.HideLocalSKUonINV Then
                        .elText = ApproLines.Item(i).CodeFForExport
                    Else
                        .elText = ApproLines.Item(i).CodeF
                    End If
                    .elCreateSibling "Title", True
                        .elText = ApproLines.Item(i).Title & " / " & ApproLines.Item(i).Author
                    .elCreateSibling "Qty", True
                        .elText = ApproLines.Item(i).Qty
                    .elCreateSibling "Price", True
                        .elText = ApproLines.Item(i).PriceF
                    .elCreateSibling "DiscountRate", True
                        .elText = ApproLines.Item(i).DiscountF
                    .elCreateSibling "Reference", True
                        .elText = ApproLines.Item(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = ApproLines.Item(i).ExtensionNetF
                    .elCreateSibling "Note", True
                        .elText = ApproLines.Item(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = ApproLines.Item(i).ExtensionNetExVATF
                    .elCreateSibling "VAT", True
                        .elText = ApproLines.Item(i).ExtensionVATOnly
                    .elCreateSibling "VATRate", True
                        .elText = ApproLines.Item(i).VATRateF
                    .navUP
            Next i
                                p 10
            .elCreateSibling "TotalQtyItems", True
                .elText = Me.TotalQtyItems
                                p 11

            .elCreateSibling "TotalNumberOfLines", True
                .elText = CStr(Me.ApproLines.Count)
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
                .elText = IIf(Me.StaffName > "", Me.StaffName, "")
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.QuotationText
            
            .elCreateSibling "VATable", True
                .elText = IIf(Me.VATable, "TRUE", "FALSE")
            .elCreateSibling "ShowVAT", True
                .elText = IIf(Me.ShowVAT, "TRUE", "FALSE")
                
            .elCreateSibling "TotalIncVAT", True
                .elText = Me.TotalNetF
            .elCreateSibling "TotalExVAT", True
                .elText = Me.TotalNetExVATF
            .elCreateSibling "TotalVAT", True
                .elText = Me.TotalVATF
    End With
                                p 12

'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "APP_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

                                p 13

    If DispatchMode = enMail Then
        'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_HTML.XSLT"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
'MsgBox "Check this"
            If fs.FileExists(strWorkingFolder & "APP_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "APP_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "APP_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage
            oTF.CloseTextFile
    End If
    
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If (Me.Customer.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
                        p 6
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        strFOFile = strWorkingFolder & "APP_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "APP_" & Me.DOCCode & ".PDF"
                        p 7
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
 '       strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        ChDir "\PBKS\Executables\FOP\"
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000

        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If
    ExportToXML = True
    
    If (Not fs.FileExists(strPDFFile)) And (DispatchMode = enView Or DispatchMode = enMail) Then
            MsgBox "The PDF file has not been produced: " & strPDFFile & ". Please contact support.", vbOKOnly, "Can't do this"
    Else
        If DispatchMode = enView Or DispatchMode = enPrint Then
            OpenFileWithApplication strPDFFile, enPDF, QuickPrint
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
End Function
Public Function ExportToSpreadsheet(bForeign As Boolean, _
        pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim s As String
Dim s2 As String
Dim lngNumberOfLines As Long

    ExportToSpreadsheet = False
    pFilename = oPC.LocalFolder & "Temp\" & DOCCode & "_" & Format(Now(), "yyyymmddHHnn") & ".xls"
    
    oTF.OpenTextFile pFilename
    oTF.WriteToTextFile "Status" & vbTab & Me.STATUS
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCode
    oTF.WriteToTextFile "DocDate" & vbTab & DocDateF
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.DefaultCompany.CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "CustomerName" & vbTab & Customer.Name
    oTF.WriteToTextFile "OurAcnowithCustomer" & vbTab & Customer.OurACnoWithClient
    oTF.WriteToTextFile "ACNO" & vbTab & Customer.AcNo
    oTF.WriteToTextFile "BankingDetails" & vbTab & Replace(oPC.Configuration.DefaultCompany.BankDetails, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile_NoLineTerminator "CustomerWithAddress" & vbTab
                If Not ApproToAddress Is Nothing Then
                    oTF.WriteToTextFile Replace(Me.ApproToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
    
    oTF.WriteToTextFile_NoLineTerminator "CustomerPhone" & vbTab
                If Not Me.ApproToAddress Is Nothing Then
                    oTF.WriteToTextFile IIf(Me.ApproToAddress.Phone > "", "Phone: " & Me.ApproToAddress.Phone, "")
                End If
    oTF.WriteToTextFile_NoLineTerminator "CustomerFax" & vbTab
                If Not ApproToAddress Is Nothing Then
                    oTF.WriteToTextFile IIf(Me.ApproToAddress.Fax > "", "Fax: " & Me.ApproToAddress.Fax, "")
                End If
            
        If Not ApproToAddress Is Nothing Then
            s = Replace(ApproToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.ApproToAddress.Phone > "", "Phone: " & Me.ApproToAddress.Phone, "") & Chr(10) & IIf(Me.ApproToAddress.Fax > "", "Fax: " & Me.ApproToAddress.Fax, "")
        Else
            s = ""
        End If
    oTF.WriteToTextFile "BillTo" & vbTab & s
            
    s = "SKU" & vbTab & "Title" & vbTab & "Qty" & vbTab & "Price" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab _
    & "Extension" & vbTab & "Note" & vbTab & "ExtensionExVat" & vbTab & "ExtensionVATOnly" & vbTab & "VATRate" & vbTab
    
    oTF.WriteToTextFile s
                           
                            
        lngNumberOfLines = 0
        For i = 1 To ApproLines.Count
                lngNumberOfLines = lngNumberOfLines + 1
                    If oPC.HideLocalSKUonINV Then
                        s = ApproLines.Item(i).CodeFForExport
                    Else
                        s = ApproLines.Item(i).CodeF
                    End If
                
                s = s & vbTab _
                & ApproLines.Item(i).Title & " / " & ApproLines.Item(i).Author
                
                s = s & vbTab & ApproLines.Item(i).Qty
                s = s & vbTab & ApproLines.Item(i).PriceF
                s = s & vbTab & ApproLines.Item(i).DiscountF
                s = s & vbTab & ApproLines.Item(i).Ref
                s = s & vbTab & ApproLines.Item(i).ExtensionNetF
                s = s & vbTab & ApproLines.Item(i).Note
                s = s & vbTab & ApproLines.Item(i).ExtensionNetExVATF
                s = s & vbTab & ApproLines.Item(i).ExtensionVATOnly
                s = s & vbTab & ApproLines.Item(i).VATRateF
                oTF.WriteToTextFile s
            Next
            
    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(Me.ApproLines.Count)
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.WriteToTextFile "StaffMember" & vbTab & IIf(Me.StaffName > "", Me.StaffName, "")
    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalNetF
    oTF.WriteToTextFile "TotalExVAT" & vbTab & TotalNetExVATF
    oTF.WriteToTextFile "TotalVAT" & vbTab & TotalVATF
    oTF.WriteToTextFile "VATable" & vbTab & IIf(Me.VATable, "TRUE", "FALSE")
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
    ErrorIn "a_QU.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub

'''Private Sub SetupDocumentTotals(bForeign As Boolean)
'''    On Error GoTo errHandler
'''Dim i, k As Integer
'''Dim bBottomOfDocument As Boolean
'''
'''    ReDim arDocTotals(16)
'''    ReDim arDocTotalsF(16)
'''
'''    jTotal = 1
'''    RunningTotal = 0
'''
''''    If Me.TotalNet = Me.TotalExtWithOutExtras(bForeign) Then  'THer are no extras
''''        addARow "Total", TotalExtWithOutExtras(bForeign), TotalExtWithOutExtrasF(bForeign), "", "+"
''''    Else
''''        addARow "subtotal", TotalExtWithOutExtras(bForeign), TotalExtWithOutExtrasF(bForeign), "", "+"
''''    End If
'''
'''    If Customer.VATable Then
'''        'Handle extras
'''            bBottomOfDocument = False
'''            For k = 1 To ApproLines.Count
'''                If ApproLines(k).BottomOfDocument = True Then
'''                    addARow ApproLines(k).Title, ApproLines(k).ForeignPrice, ApproLines(k).PriceF, "", "+"
'''                    bBottomOfDocument = True
'''                End If
'''            Next k
'''            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "GRAND"
'''            addARow "(includes VAT at " & oPC.Configuration.VATRateF, TotalVATRaw(bForeign), TotalVATF(bForeign), "", " "
'''    Else
'''            bBottomOfDocument = False
'''            For k = 1 To InvoiceLines.Count
'''                If InvoiceLines(k).BottomOfDocument = True And InvoiceLines(k).VATRate <> 0 Then
'''                    addARow InvoiceLines(k).Title, InvoiceLines(k).ForeignPrice, InvoiceLines(k).PriceF(bForeign), "", "+"
'''                    bBottomOfDocument = True
'''                End If
'''            Next k
'''            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "SUB"
'''            If ShowVAT Then
'''                addARow "less VAT at " & oPC.Configuration.VATRateF, TotalVATRaw(bForeign), TotalVATF(bForeign), "SUB", "-"
'''            End If
'''            For k = 1 To InvoiceLines.Count
'''                If InvoiceLines(k).BottomOfDocument = True And InvoiceLines(k).VATRate = 0 Then
'''                    addARow InvoiceLines(k).Title, InvoiceLines(k).ForeignPrice, InvoiceLines(k).PriceF(bForeign), "", "+"
'''                    bBottomOfDocument = True
'''                End If
'''            Next k
'''            If bBottomOfDocument = True Or ShowVAT = True Then arDocTotals(jTotal - 1).TotalType = "GRAND"
'''    End If
'''
''''    TransferToFinalTable
'''    k = 0
'''    For i = 1 To UBound(arDocTotals)
'''        If arDocTotals(i).Description = "" Then Exit For
'''        k = k + 1
'''        arDocTotalsF(k).Description = arDocTotals(i).Description
'''        arDocTotalsF(k).AmtFormatted = arDocTotals(i).AmtFormatted
'''        If arDocTotals(i).TotalType > "" Then
'''            k = k + 1
'''            arDocTotalsF(k).Description = IIf(arDocTotals(i).TotalType = "GRAND", "Total", "subtotal")
'''            arDocTotalsF(k).val = CStr(arDocTotals(i).RunningTotal)
'''            If bForeign Then
'''                arDocTotalsF(k).AmtFormatted = Format(arDocTotals(i).RunningTotal / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
'''            Else
'''                arDocTotalsF(k).AmtFormatted = Format(arDocTotals(i).RunningTotal / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'''            End If
'''        End If
'''    Next i
'''
'''    Exit Sub
'''errHandler:
'''    If ErrMustStop Then Debug.Assert False: Resume
'''    ErrorIn "a_Invoice.SetupDocumentTotals(bForeign)", bForeign
'''End Sub
'''
'''Private Sub addARow(pDescription, Amt As Long, AmtF As String, pTotalType As String, Optional Sign As String)
'''    On Error GoTo errHandler
'''    arDocTotals(jTotal).Description = pDescription
'''    arDocTotals(jTotal).val = Amt
'''    arDocTotals(jTotal).AmtFormatted = AmtF & IIf(Left(pDescription, 1) = "(", ")", "")
'''    arDocTotals(jTotal).TotalType = pTotalType
'''    If Sign = "-" Then
'''        RunningTotal = RunningTotal - Amt
'''    ElseIf Sign = "+" Then
'''        RunningTotal = RunningTotal + Amt
'''    End If
'''    arDocTotals(jTotal).RunningTotal = RunningTotal
'''    jTotal = jTotal + 1
'''    Exit Sub
'''errHandler:
'''    If ErrMustStop Then Debug.Assert False: Resume
'''    ErrorIn "a_Invoice.addARow(pDescription,Amt,AmtF,pTotalType)", Array(pDescription, Amt, _
'''         AmtF, pTotalType)
'''End Sub


