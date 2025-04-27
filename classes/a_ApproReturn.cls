VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(pMsg As String)
Private colClassors As Collection
Private mudtProps As APPRProps
Private mcolStack As Collection
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private mColAppRetLines As ch_APPRL
Dim mOriginalStatus As enStatus
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Dim oPCL As z_PrintClient
Const CDOCCODE = "AR"

Public Property Get Customer() As a_Customer
    Set Customer = mCustomer
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
  Set mColAppRetLines = New ch_APPRL
    mColAppRetLines.component Me
    Set mCustomer = New a_Customer
  Set mcolStack = New Collection
    mOriginalStatus = stInProcess
    LoadClassorsCollection
    mudtProps.COMPID = oPC.Configuration.DefaultCompany.ID
    mudtProps.Status = stInProcess
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set mobjValid = Nothing
    Set mColAppRetLines = Nothing
    Set mCustomer = Nothing
    Set mcolStack = Nothing
    Set colClassors = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.mobjValid_Status(pMsg)", pMsg
End Sub
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As APPRData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As APPRData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "ARLS", mColAppRetLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.GetSuperState"
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
    mColAppRetLines.SetSuperState .ReadProperty("ARLS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.SetSuperState(buffer)", buffer
End Sub
Public Function SetCustomer(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    mudtProps.TPID = pTPID
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.SetCustomer(pTPID)", pTPID
End Function

Public Sub BeginEdit()
    On Error GoTo errHandler
  mColAppRetLines.BeginEdit
  ' save object state
  mcolStack.Add GetState
  Debug.Print "BeginEdit " & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mColAppRetLines.CancelEdit
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
  Debug.Print "CancelEdit " & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
Dim objPersist As a_APPR_P
  If mcolStack.Count = 0 Then Err.Raise 445
    If DOCCode = "" Then
        DOCCode = GetAPPRNumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
  mColAppRetLines.ApplyEdit
  Set objPersist = New a_APPR_P
  If mudtProps.IsDeleted Then
    ' code to delete the object's data goes here
    objPersist.DeleteObject mudtProps.TRID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    ' save object to database if appropriate
    ' save object state
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
    ErrorIn "a_APPR.ApplyEdit"
End Sub

Public Sub Load(ApproRetID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_APPR_P
  
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    ' code to load the object goes here
    Set objPersist = New a_APPR_P
    SetSuperState objPersist.Fetch(ApproRetID)
    Set objPersist = Nothing
    SetCustomer Me.TPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.Load(ApproRetID,ReadOnly)", Array(ApproRetID, ReadOnly)
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mColAppRetLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IsValid"
End Property

Public Property Get TPNAME() As String
    On Error GoTo errHandler
    TPNAME = FNS(mudtProps.TPNAME)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.TPName"
End Property

Public Property Let TPNAME(Value As String)
    On Error GoTo errHandler
    mudtProps.TPNAME = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.TPName(Value)", Value
End Property

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
    ErrorIn "a_APPR.StatusF"
End Property
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.Status"
End Property
Public Property Let Memo(val As String)
On Error Resume Next
    mudtProps.Memo = Trim(val)
End Property

Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.Memo"
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
    

Public Property Get IssDateF() As String
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IssDateF"
End Property

Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IssDate"
End Property
    
Public Property Let IssDate(Value As Date)
    On Error GoTo errHandler
      If IsDate(Value) Then
        mudtProps.CaptureDate = CDate(Value)
        mudtProps.IsDirty = True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IssDate(Value)", Value
End Property
    
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.DocDateF"
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.DocDate"
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
    ErrorIn "a_APPR.DocDate(Value)", Value
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.DocCode"
End Property

Public Property Let DOCCode(Value As String)
    On Error GoTo errHandler
    mudtProps.DOCCode = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.DocCode(Value)", Value
End Property

'Public Property Get APPID() As Long
'    APPID = mudtProps.APPID
'End Property
'
'Public Property Let APPID(Value As Long)
'    mudtProps.APPID = Value
'End Property

Public Property Get APPRLines() As ch_APPRL
  Set APPRLines = mColAppRetLines
End Property

Public Property Get DisplayAPPR() As d_APPR
    On Error GoTo errHandler
    Set DisplayAPPR = New d_APPR
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.DisplayAPPR"
End Property
'Public Property Get TPACCNum() As String
'    TPACCNum = FNS(mudtProps.TPACCNum)
'End Property
'Public Property Let TPACCNum(val As String)
'    mudtProps.TPACCNum = val
'  mudtProps.IsDirty = True
'End Property

Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.TPID"
End Property

Public Property Let TPID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TPID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.TPID(Value)", Value
End Property

Public Property Get COMPID() As Long
    On Error GoTo errHandler
    COMPID = mudtProps.COMPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.COMPID"
End Property

Public Property Let COMPID(Value As Long)
    On Error GoTo errHandler
    mudtProps.COMPID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.COMPID(Value)", Value
End Property

Public Function GetAppRetNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String

    Set oCode = New z_Code
    
    GetAppRetNumber = oCode.GetNextCode("AppRet", strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetAppRetNumber = strDefaultCo & "AR" & GetAppRetNumber
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.GetAppRetNumber(strTmp)", strTmp
End Function

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.TRID"
End Property

Public Property Let TRID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TRID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.TRID(Value)", Value
End Property
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.StaffID"
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
    ErrorIn "a_APPR.StaffName"
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
    ErrorIn "a_APPR.StaffNameB"
End Property

Public Property Let StaffID(Value As Long)
    On Error GoTo errHandler
    mudtProps.StaffID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.StaffID(Value)", Value
End Property

Public Function GetAPPRNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strWork As String

    Set oCode = New z_Code
    
    strWork = oCode.GetNextCode(enApproReturn, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetAPPRNumber = strDefaultCo & "AR" & strWork
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.GetAPPRNumber(strTmp)", strTmp
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
    ErrorIn "a_APPR.TranslateErrors(pRawors)", pRawors
End Function
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.IsEditing"
End Property
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.LoadClassorsCollection"
End Sub

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
    ErrorIn "a_APPR.GetSTatus"
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.SetStatus(val)", val
End Sub

Public Function PrintAPPR() As Boolean
    On Error GoTo errHandler
Dim iTotalLinesCOunt As Integer
Dim oDC As a_DocumentControl

    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintAPPR = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("APR_" & Me.DOCCode) = False Then
           PrintAPPR = False
           Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
        End With
        PrintAPPR = True
    End If
    
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt APPRLines.Count
    iTotalLinesCOunt = 0
    oPCL.LineTotalCount iTotalLinesCOunt 'This must be written before the TOTAL tags for the Printserver requirements
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "CUSTOMER", Me.Customer.NameAndCode(25)
    oPCL.WriteTag "DATE", Me.DocDateF
    oPCL.WriteTag "STORECODE", oPC.Configuration.DefaultStore.code
    oPCL.WriteTag "MEMO", PackText(Memo)
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")

    FillTableRows
    oPCL.FinishRequest

EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.PrintAPPR"
End Function
Private Function FillTableRows()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To APPRLines.Count
        With APPRLines(i)
            oPCL.APPRLSend .CodeF, .Qty, .Title, .QtyIssued & " (" & .QtyReturned & ")", .ApproCode & " (" & .ApproDateF & ")", PackText(.Note)
        End With
    Next i
EXIT_Handler:
'Err_Handler:
'    MsgBox error
'    GoTo EXIT_Handler
   ' Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.FillTableRows"
End Function


Public Function Post() As String
    On Error GoTo errHandler
Dim strResult As String
Dim oSM As z_StockManager
Dim strPos As String

strPos = "Pos 1"
'    oPC.COShort.BeginTrans
strPos = "Pos 2"
    Set oSM = New z_StockManager
strPos = "Pos 3"
    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
strPos = "Pos 4"
            Select Case Status
            Case stVOID
       ''         SetStatus stVOID
            Case stInProcess
                Post = "In Process"
            Case stCOMPLETE
strPos = "Pos 4.1"
                oSM.IssueAPPR Me
            Case stISSUED
strPos = "Pos 4.2"
                oSM.IssueAPPR Me
              '  SetStatus stCOMPLETE
            Case stPROFORMA
                oSM.IssueAPPR Me
        '        SetStatus stPROFORMA
            End Select
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    Case stPROFORMA      'Issued PROFORMA
            Select Case Status
            Case stVOID
                oSM.CancelAPPR Me
        '        SetStatus stVOID
            Case stInProcess
                Post = "OR: was proforma, can't now be inprocess."
            Case stCOMPLETE
                oSM.IssueAPPR Me
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    Set oSM = Nothing
strPos = "Pos 5"
'    oPC.COShort.CommitTrans
strPos = "Pos 6"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.Post", , EA_RERAISE + EA_ADVANCED, oPC.COShort, "mOriginalStatus,Status,strPos", Array(mOriginalStatus, Status, strPos)
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.SetDirty(pVal)", pVal
End Sub
Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPR.VoidDocument"
End Function
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
    ErrorIn "a_APPR.Reload"
End Sub

