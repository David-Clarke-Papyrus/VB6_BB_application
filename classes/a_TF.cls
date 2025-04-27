VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_TF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient

Private mExtExVAT As Long
Private mExtLessDiscExVAT As Long
Private mDiscExVAT As Long
Private mQtyTotal As Long
Private mTotalCost As Long
Private tlOperators As z_TextList
Private mudtProps As TFProps
Private mcolStack As Collection
Private WithEvents mColTFLines As ch_TFL
Attribute mColTFLines.VB_VarHelpID = -1
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Event Valid(pMsg As String)
Event TotalChange(Total As String, QtyTotal As String)
Event Dirty(pVal As Boolean)
Const CDOCCODE = "TF"
Private xMLDoc As ujXML
Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property

Public Function CalculateTotal()
Dim bLineDiscountApplied As Boolean
Dim ol As a_TFL
    bLineDiscountApplied = False
    
    mExtExVAT = 0
    mExtLessDiscExVAT = 0
    mDiscExVAT = 0
    mQtyTotal = 0
    mTotalCost = 0
    For Each ol In mColTFLines
            ol.RecalculateLine
            mExtExVAT = mExtExVAT + ol.ExtExVAT
            mDiscExVAT = mDiscExVAT + ol.DiscExVAT
            mExtLessDiscExVAT = mExtLessDiscExVAT + ol.ExtLessDiscExVAT
            mQtyTotal = mQtyTotal + ol.Qty
            mTotalCost = mTotalCost + ol.LineCost
    Next
    mudtProps.ExtExVAT = mExtExVAT
    mudtProps.ExtLessDiscExVAT = mExtLessDiscExVAT
    mudtProps.DiscExVAT = mDiscExVAT
    mudtProps.TotalQtyItems = mQtyTotal
    mudtProps.TotalCostExVAT = mTotalCost
    ValidateObject "TOTAL"
    
    RaiseEvent TotalChange(Me.ExtLessDiscExVATF, Me.TotalQtyItemsF)

End Function

Public Property Get TotalCostExVAT() As Long
    TotalCostExVAT = mudtProps.TotalCostExVAT
End Property
Public Property Get TotalCostExVATF() As String
        TotalCostExVATF = Format(CDbl(mudtProps.TotalCostExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get TotalQtyItems() As Long
    TotalQtyItems = mudtProps.TotalQtyItems
End Property
Public Property Get TotalQtyItemsF() As String
        TotalQtyItemsF = Format(CDbl(mudtProps.TotalQtyItems), "###,##0")
End Property

Public Property Get ExtExVAT() As Long
    ExtExVAT = mudtProps.ExtExVAT
End Property
Public Property Get ExtExVATF() As String
        ExtExVATF = Format(CDbl(mudtProps.ExtExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get ExtLessDiscExVAT() As Long
    ExtLessDiscExVAT = mudtProps.ExtLessDiscExVAT
End Property
Public Property Get ExtLessDiscExVATF() As String
        ExtLessDiscExVATF = Format(CDbl((ExtLessDiscExVAT) / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get DiscExVAT() As Long
    DiscExVAT = mudtProps.DiscExVAT
End Property
Public Property Get DiscExVATF() As String
        DiscExVATF = Format(CDbl(DiscExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

'Public Property Get TotalPayable() As String
'        TotalPayable = mudtProps.Payable
'End Property
'Public Property Get TotalPayableF() As String
'        TotalPayableF = Format(CDbl(mudtProps.Payable / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'
'Public Property Get TotalPayableExVAT() As Long
'        TotalPayableExVAT = mudtProps.PayableExVAT
'End Property
'Public Property Get TotalPayableExVATF() As String
'        TotalPayableExVATF = Format(CDbl(mudtProps.PayableExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'
'Public Property Get TotalVAT() As Long
'    TotalVAT = TotalPayable - TotalPayableExVAT
'End Property
'Public Property Get TotalVATF() As String
'    TotalVATF = Format(CDbl(TotalVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'End Property

Public Function DisplayTotals(pLabels As String, pValues As String)
Dim strLabels As String
Dim strValues As String
Dim i As Integer


    strLabels = ""
    strValues = ""
'    If oPC.TransferIsExVAT = True Then
        strLabels = strLabels & Chr(13) & "Total"
        strValues = strValues & Chr(13) & Me.ExtLessDiscExVATF
'    Else
'        strLabels = strLabels & Chr(13) & "subtotal" & Chr(13) & "Includes V.A.T. of" & Chr(13) & "Nett value"
'        strValues = strValues & Chr(13) & Me.TotalPayableF & Chr(13) & Me.TotalVATF & Chr(13) & Me.TotalPayableExVATF
'    End If
    pLabels = strLabels
    pValues = strValues
End Function
Public Sub Reload()
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID
End Sub

Public Sub VoidDocument()
Dim strResult As String
    If Me.STATUS <> stVOID Then
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    End If
End Sub
Public Property Get STATUS() As enStatus
    STATUS = mudtProps.STATUS
End Property

Private Sub Class_Initialize()
    mOriginalStatus = stInProcess
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set mColTFLines = New ch_TFL
    mColTFLines.component Me
    Set mcolStack = New Collection
    LoadClassorsCollection
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    mudtProps.STATUS = 2
    Me.InOut = "IN"
End Sub

Private Sub Class_Terminate()
    Set mobjValid = Nothing
    Set mColTFLines = Nothing
    Set mcolStack = Nothing
    Set colClassors = Nothing
End Sub

Private Sub mColTFLines_rowsChange()
    RaiseEvent TotalChange(Me.ExtLessDiscExVATF, Me.TotalQtyItemsF)
End Sub

Private Sub mobjValid_BrokenRule(ors As String)
  RaiseEvent Valid(ors)
End Sub

Private Sub mobjValid_NoBrokenRules()
  RaiseEvent Valid("")
End Sub
Private Function GetState() As String
  Dim udtData As TFData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
End Function
Private Sub SetState(Buffer As String)
  Dim udtData As TFData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData
End Sub
Public Function GetSuperState() As String
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "TFLS", mColTFLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function
Public Sub SetSuperState(Buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = Buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mColTFLines.SetSuperState .ReadProperty("TFLS")
  End With
  Set objPB = Nothing
End Sub

Public Property Get StatusF() As String
    If mudtProps.STATUS = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.STATUS = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.STATUS = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.STATUS = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
End Property

Public Property Let STATUS(val As Long)
    mudtProps.STATUS = val
    SetDirty True
End Property
Public Sub SetStatus(val As enStatus)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.STATUS = val
    SetDirty True
End Sub


Public Sub BeginEdit()
  mColTFLines.BeginEdit
  mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mColTFLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Public Sub ApplyEdit()
Dim objPersist As a_TF_P
  If mcolStack.Count = 0 Then Err.Raise 445
    If DOCCode = "" Then
        DOCCode = GetTFNumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now

  mColTFLines.ApplyEdit
  Set objPersist = New a_TF_P
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject mudtProps.TRID
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Me.STATUS = stISSUED Then
        If Not IsValid Then Err.Raise 445
    End If
    mcolStack.Remove mcolStack.Count
    SetSuperState objPersist.Save(GetSuperState)
    mudtProps.IsNew = False
  Else
    mcolStack.Remove mcolStack.Count
  End If
  Set objPersist = Nothing
  mudtProps.IsDirty = False
End Sub

Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = mcolStack.Count > 0
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty Or mColTFLines.IsDirty
End Property

Public Property Get IsValid() As Boolean
    IsValid = (mobjValid.Count = 0)
End Property
Public Property Get IsAuto() As Boolean
    IsAuto = mudtProps.Auto
End Property

Public Function GetSTatus()
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
End Function

Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub
Public Sub Load(pTFID As Long)
Dim oPersist As a_TF_P
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    Set oPersist = New a_TF_P
    SetSuperState oPersist.Fetch(pTFID)
    Set oPersist = Nothing
    mOriginalStatus = Me.STATUS
    CalculateTotal
    mudtProps.IsNew = False
End Sub
'Public Sub Delete()
'    If Not flgEditing Then Err.Raise 445
'    flgDeleted = True
'End Sub
Private Sub DeleteObject(lngResult As Long, ID As Long)
On Error GoTo ERR_Handler

  Dim strSQL As String
  Dim lngRecs As Long
  
    strSQL = "DELETE * FROM tblTA WHERE Transaction_ID=" & ID
    oPC.COShort.execute strSQL
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    If ((Err = DB_EXISTS) Or (Err = DB_MISSING) Or (Err = DB_REF) Or (Err = DB_FIELDCONSTRAINT)) Then
        lngResult = Err
    Else
        tmp = Err
        tmpor = Error
        Err.Raise tmp
    End If
    GoTo EXIT_Handler

End Sub
Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Get DestID() As Long
    DestID = mudtProps.DestID
End Property
Public Property Let DestID(val As Long)
    If Len(val) > Len(mudtProps.DestID) Then Err.Raise vbObjectError + 1001, "String value too long"
    mobjValid.RuleBroken "TPID", (Len(Trim$(val)) = 0)
    mudtProps.DestID = val
    SetDirty True
End Property
Public Property Get StaffID() As Long
    StaffID = mudtProps.StaffID
End Property
Public Property Let StaffID(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
End Property
Public Property Get StaffName() As String
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).StaffName
End Property
Public Property Get StaffNameB() As String
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
End Property

Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Get DestinationName() As String
    DestinationName = FNS(mudtProps.DestinationName)
End Property
Public Property Let DestinationName(val As String)
    mudtProps.DestinationName = val
End Property
Public Property Let DOCCode(val As String)
    mudtProps.DOCCode = val
    SetDirty True
End Property
Public Property Get InOut() As String
    InOut = FNS(mudtProps.InOut)
End Property
Public Property Let InOut(val As String)
    mobjValid.RuleBroken "INOUT", Not (val = "IN" Or val = "OUT")
    mudtProps.InOut = val
    SetDirty True
End Property
Public Property Let DOCDate(val As Date)
    
    mobjValid.RuleBroken "DOCDATE", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.DOCDate = val
    SetDirty True
End Property
Public Property Get DOCDate() As Date
    DOCDate = mudtProps.DOCDate
End Property
Public Property Get DocDateF() As String
    If CLng(mudtProps.DOCDate) = 0 Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    End If
End Property
Public Property Let IssDate(val As Date)
    
    mobjValid.RuleBroken "CaptureDate", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.CaptureDate = val
    SetDirty True
End Property
Public Property Get IssDate() As Date
    IssDate = mudtProps.CaptureDate
End Property
Public Property Get IssDateF() As String
    If CLng(mudtProps.CaptureDate) = 0 Then
        IssDateF = ""
    Else
        IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    End If
End Property
Public Property Get Operators() As z_TextList
    Set Operators = tlOperators
End Property

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
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Invalid type", "INOUT"
    colClassors.Add "Batch control value mismatch", "BATCHTOTAL"
    colClassors.Add "Batch control quantity mismatch", "BATCHQTYTOTAL"
End Sub
Public Function GetTFNumber(strTmp As String) As String
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strWork As String

    Set oCode = New z_Code
    
    strWork = oCode.GetNextCode(enTransfer, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetTFNumber = strDefaultCo & "T" & strWork
    
    Set oCode = Nothing
End Function

Private Sub SetDirty(pVal As Boolean)
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
End Sub
Public Property Get TFLines() As ch_TFL
    Set TFLines = mColTFLines
End Property

Public Function Post() As String
Dim oSM As z_StockManager
On Error GoTo H
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
 '   oPC.COShort.BeginTrans
    Me.ApplyEdit
    Set oSM = New z_StockManager
    
    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
            Select Case STATUS
            Case stVOID
            Case stInProcess
                Post = "In Process"
            Case stCOMPLETE, stISSUED
                oSM.PostTransfer Me.TRID, stCOMPLETE, False
            End Select
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    End Select
    Set oSM = Nothing
  '  oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox "a_TF:Post: " & Error
    oPC.COShort.RollbackTrans
    Exit Function
    Resume
End Function
Public Sub PasteLine(PID As String, Qty As Long, Price As Long, DiscountRate As Double, VATRate As Double)
Dim oSM As New z_StockManager

    oSM.PasteTransferLine PID, Qty, Price, DiscountRate, VATRate, mudtProps.TRID
    
End Sub

Public Function PrintTF() As Boolean
Dim iTotalLinesCOunt As Integer
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl
    On Error GoTo ERR_Handler
    CalculateTotal
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintTF = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("TFR_" & Me.DOCCode) = False Then
            PrintTF = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
        End With
        PrintTF = True
    End If
    If Me.STATUS = stInProcess Or STATUS = stCANCELLED Or STATUS = stVOID Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt TFLines.Count
    iTotalLinesCOunt = 3
    oPCL.LineTotalCount iTotalLinesCOunt 'This must be written before the TOTAL tags for the Printserver requirements
    
    oPCL.WriteTag "INOUT", Me.InOut
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "OURADDRESS", oPC.Configuration.DefaultCompany.StreetAddressDelimited
    oPCL.WriteTag "DATE", Me.DocDateF
    oPCL.WriteTag "STORECODE", oPC.Configuration.DefaultStore.code
    oPCL.WriteTag "MEMO", PackText(Memo)
    oPCL.WriteTag "DELTO", IIf(InOut = "IN", "FROM: ", "TO: ") & oPC.Configuration.Stores.FindStoreByID(DestID).Description
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    FillTableRows
'    If oPC.TransferIsExVAT = True Then
        oPCL.WriteTag "TOTAL1", Me.ExtExVATF  'TotalLessDiscExtF
        oPCL.WriteTag "TOTAL2", Me.DiscExVATF
        oPCL.WriteTag "TOTAL3", Me.ExtLessDiscExVATF
        oPCL.WriteTag "TOTALDESCRIPTION", PackText("Total Ex VAT" & vbCrLf & "Discount Ex VAT" & vbCrLf & "Nett value Ex VAT")
'    Else
'        oPCL.WriteTag "TOTAL1", Me.TotalPayableF  'TotalLessDiscExtF
'        oPCL.WriteTag "TOTAL2", Me.TotalVATF
'        oPCL.WriteTag "TOTAL3", Me.TotalPayableExVATF
'        oPCL.WriteTag "TOTALDESCRIPTION", PackText("subtotal" & vbCrLf & "Includes V.A.T. of" & vbCrLf & "Nett value")
'    End If
    oPCL.FinishRequest

EXIT_Handler:
    Exit Function
ERR_Handler:
    oPCL.FinishRequest
    Select Case Err
    Case 5941
        MsgBox "Bookmark missing"
        Resume Next
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select
End Function
Private Function FillTableRows()
Dim i As Integer
On Error GoTo ERR_Handler
    For i = 1 To TFLines.Count
        With TFLines(i)
          '  If oPC.TransferIsExVAT = True Then
                oPCL.TFLSend .CodeF, .Qty, .PriceF, .DiscountF, .ExtLessDiscExVATF, .ExtLessDiscExVATF, .SpecialVATMsg & .Title, 0
          '  Else
          '      oPCL.TFLSend .CodeF, .Qty, .PriceF, .DiscountF, .ExtPayableF, "", .SpecialVATMsg & .Title, 0
          '  End If
        End With
    Next i
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
   ' Resume
End Function

Public Function SetMemo(val As String) As Boolean
   ' If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Memo) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Memo Then
        mudtProps.Memo = val
        SetDirty True
    End If
    SetMemo = True
End Function
Public Property Get Memo() As String
    Memo = FNS(mudtProps.Memo)
End Property

Public Property Let Memo(val As String)
    On Error Resume Next
    mudtProps.Memo = Trim(val)
End Property

Public Function SetBatchQtyTotal(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long

        lngTemp = mudtProps.BatchQtyTotal
        bOK = SetField_LONG(mudtProps.BatchQtyTotal, val, "BATCHQTYTOTAL", mcolStack.Count)
        bOK = ValidateObject("BATCHQTYTOTAL")
        SetBatchQtyTotal = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.SetBatchQtyTotal(val)", val
End Function

Public Property Get BatchQtyTotal() As Long
    On Error GoTo errHandler
    BatchQtyTotal = mudtProps.BatchQtyTotal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.BatchQtyTotal"
End Property
Public Property Get BatchQtyTotalF() As String
    On Error GoTo errHandler
    BatchQtyTotalF = CStr(mudtProps.BatchQtyTotal)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.BatchQtyTotalF"
End Property

Public Property Get SendersDocRef() As String
    On Error GoTo errHandler
    SendersDocRef = FNS(mudtProps.SendersDocRef)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.SendersDocRef"
End Property

Public Property Let SendersDocRef(Value As String)
    On Error GoTo errHandler
    mudtProps.SendersDocRef = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.SendersDocRef(Value)", Value
End Property
Public Property Get SendersDocDate() As Date
    On Error GoTo errHandler
    SendersDocDate = mudtProps.SendersDocDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.SendersDocDate"
End Property
Public Property Get SendersDocDateF() As String
    On Error GoTo errHandler
    If CLng(mudtProps.SendersDocDate) = 0 Then
        SendersDocDateF = ""
    Else
        SendersDocDateF = Format(mudtProps.SendersDocDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.SendersDocDateF"
End Property

Public Property Let SendersDocDate(Value As Date)
    On Error GoTo errHandler
    mudtProps.SendersDocDate = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.SendersDocDate(Value)", Value
End Property
Friend Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If mudtProps.DestID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If
    If Me.InOut = "IN" And Me.IsAuto = False Then
        If (mudtProps.BatchTotal > (Me.ExtLessDiscExVAT * (1 + oPC.SupplierInvoiceTolerance))) _
            Or (mudtProps.BatchTotal < (Me.ExtLessDiscExVAT * (1 - oPC.SupplierInvoiceTolerance))) Then
            mobjValid.BreakRule "BATCHTOTAL", True
        Else
            mobjValid.BreakRule "BATCHTOTAL", False
        End If
        If mudtProps.BatchQtyTotal <> (mudtProps.TotalQtyItems) Then
            mobjValid.BreakRule "BATCHQTYTOTAL", True
        Else
            mobjValid.BreakRule "BATCHQTYTOTAL", False
        End If
    Else
        mobjValid.BreakRule "BATCHTOTAL", False
        mobjValid.BreakRule "BATCHQTYTOTAL", False
    End If
    mobjValid.GetSTatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFR.ValidateObject(pFld)", pFld
End Function

Public Property Get BatchTotal() As Long
    On Error GoTo errHandler
    BatchTotal = mudtProps.BatchTotal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF.BatchTotal"
End Property
Public Property Get BatchTotalF() As String
    On Error GoTo errHandler
    BatchTotalF = Format(mudtProps.BatchTotal / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF.BatchTotalF"
End Property
Public Function SetBatchTotal(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long

        lngTemp = mudtProps.BatchTotal
        bOK = SetField_LONG(mudtProps.BatchTotal, val, "BATCHTOTAL", mcolStack.Count)
        bOK = ValidateObject("BATCHTOTAL")
        SetBatchTotal = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF.SetBatchTotal(val)", val
End Function

Public Function ExportToXML(DispatchMode As enTransmitType, _
            Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, Optional pQtyCopies As Integer, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
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
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim oDC As a_DocumentControl
Dim strExecutable As String
Dim cnt As Integer


    If DispatchMode = enPrint Then
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\Printing\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\Printing\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Then
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\PDF\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\PDF\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
        strDispatchMethod = ""
    ElseIf DispatchMode = enXML Then
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\XML\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\XML\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\XML\"
        strDispatchMethod = ""
    Else
        If Not fs.FolderExists(oPC.SharedFolderRoot & "\Printing\") Then
            fs.CreateFolder oPC.SharedFolderRoot & "\Printing\"
        End If
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If
                        p 1
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
        ExportToXML = False
        Exit Function
    End If
                        p 2
    If pQtyCopies < 1 Then pQtyCopies = 1

    Set xMLDoc = New ujXML
                        p 21
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "TF_DOC"
            .chCreate "MessageType"
                .elText = "TF"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
                        p 23
            .elCreateSibling "TemplateName"
                .elText = "TF_DOC"
            .elCreateSibling "SendersEmail"
                If oPC.EmailFrom > "" Then
                    .elText = oPC.EmailFrom
                Else
                    .elText = Me.SendersEmail
                End If
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
            .elCreateSibling "DocDate", True
                .elText = Me.DOCDate
            .elCreateSibling "OurCompany", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "OurAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "SendingStoreName", True
                .elText = oPC.Configuration.DefaultStore.Description
            .elCreateSibling "DestinationStoreName", True
                .elText = Me.DestinationName
            .elCreateSibling "INOUT", True
                .elText = Me.InOut
'
            For i = 1 To Me.TFLines.Count
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    .elText = TFLines(i).CodeF
                    .elCreateSibling "EAN", True
                       .elText = TFLines(i).EAN
                    .elCreateSibling "Code", True
                       .elText = TFLines(i).code
                    .elCreateSibling "Title", True
                       .elText = TFLines(i).Title
                    .elCreateSibling "Binding", True
                       .elText = TFLines(i).BindingCode
                    .elCreateSibling "Cost", True
                       .elText = TFLines(i).CostF
                    .elCreateSibling "Author", True
                       .elText = TFLines(i).Author
                    .elCreateSibling "Qty", True
                        .elText = TFLines(i).Qty
                    .elCreateSibling "Price", True
                        .elText = TFLines(i).PriceF
                    .elCreateSibling "PriceRaw", True
                        .elText = TFLines(i).Price
                    .elCreateSibling "DiscountRate", True
                        .elText = TFLines(i).DiscountF
                    .elCreateSibling "DiscountRateRaw", True
                        .elText = TFLines(i).Discount
                    .elCreateSibling "Extension", True
                        .elText = TFLines(i).ExtLessDiscF
                    .elCreateSibling "LineCost", True
                        .elText = TFLines(i).LineCostF
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = TFLines(i).ExtLessDiscExVATF
                    .elCreateSibling "Note", True
                        .elText = TFLines(i).Note
                    .navUP
            Next i
                            p 25
            .elCreateSibling "TotalQtyItems", True
                .elText = Me.TotalQtyItemsF
            .elCreateSibling "ExtExVAT", True
                .elText = Me.ExtExVATF
            .elCreateSibling "DiscExVAT", True
                .elText = Me.DiscExVATF
            .elCreateSibling "ExtLessDiscExVAT", True
                .elText = Me.ExtLessDiscExVATF
            .elCreateSibling "TotalCostExVAT", True
                .elText = Me.TotalCostExVATF
            .elCreateSibling "VATable", True
                .elText = "TRUE"
            .elCreateSibling "ShowVAT", True
                .elText = "TRUE"
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
                If oPC.Configuration.Staff.FindStaffByID(Me.StaffID) Is Nothing Then
                    .elText = ""
                Else
                    .elText = oPC.Configuration.Staff.FindStaffByID(Me.StaffID).StaffName
                End If
            '    .elText = IIf(Me.Signature > "", Me.Signature(True), "")
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.OrderText
    End With
                            p 26
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "TF_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With
                            p 27
    If DispatchMode = enMail Then
        'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\TF_DOC_HTML.xslt"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
        
            If fs.FileExists(strWorkingFolder & "TF_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "TF_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "TF_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage 'xMLDoc.docObject.transformNode(objXSL)
            oTF.CloseTextFile
    End If
    
                        p 27
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\TF_DOC_FO.xsl"
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
        
        strFOFile = strWorkingFolder & "TF_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "TF_" & Me.DOCCode & ".PDF"
                                p 7
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
       ''' strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
 '       strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        ChDir "\PBKS\Executables\FOP\"
     strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
       ' strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
        cnt = 0
        Do While Not fs.FileExists(strPDFFile) And cnt < 40
            MsgWaitObj 1000
            cnt = cnt + 1
        Loop
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If
                        p 8
'Log the dispatch
'  '  Set oSM = New z_StockManager
'    Set oFSO = New FileSystemObject
'    pFilename = strPDFFile
'
'    If DispatchMode <> enMail And DispatchMode <> enView Then
'        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
'        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
'    End If
'    Set oSM = Nothing
'    Set oFSO = Nothing
                        p 9
    oTF.CloseTextFile
    ExportToXML = True
    
    '
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
    ErrorIn "a_CO.ExportToXML(DispatchMode,pDestinationEmailAddress,pWholeMessage,pQtyCopies)", _
         Array(DispatchMode, pDestinationEmailAddress, pWholeMessage, pQtyCopies)
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
    oTF.WriteToTextFile "DocDate" & vbTab & DOCDate
    oTF.WriteToTextFile "OurCompany" & vbTab & oPC.Configuration.DefaultCompany.CompanyName
    oTF.WriteToTextFile "OurAddress" & vbTab & Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "SendingStoreName" & vbTab & oPC.Configuration.DefaultStore.Description
    oTF.WriteToTextFile "DestinationStoreName" & vbTab & DestinationName
    oTF.WriteToTextFile "INOUT" & vbTab & InOut

    s = "SKU" & vbTab & "EAN" & vbTab & "Code" & vbTab & "Title" & vbTab & "Format" & vbTab & "Author" & vbTab & "Cost" & "Qty" & vbTab & "Price" & vbTab & "PriceRaw" & vbTab & "DiscountRate" & vbTab & "DiscountRateRaw" & vbTab _
    & "Extension" & vbTab & "LineCost" & vbTab & "ExtensionExVAT" & vbTab & "Note"
    
    oTF.WriteToTextFile s
               
                
  lngNumberOfLines = 0
  For i = 1 To TFLines.Count
    s = ""
    lngNumberOfLines = lngNumberOfLines + 1
    s = TFLines(i).CodeF
    s = s & vbTab & TFLines(i).EAN
    s = s & vbTab & TFLines(i).code
    s = s & vbTab & TFLines(i).Title
    s = s & vbTab & TFLines(i).BindingCode
    s = s & vbTab & TFLines(i).Author
    s = s & vbTab & TFLines(i).CostF
    s = s & vbTab & TFLines(i).Qty
    s = s & vbTab & TFLines(i).PriceF
    s = s & vbTab & TFLines(i).Price
    s = s & vbTab & TFLines(i).Discount
    s = s & vbTab & TFLines(i).ExtLessDiscF
    s = s & vbTab & TFLines(i).LineCostF
    s = s & vbTab & TFLines(i).ExtLessDiscExVATF
    s = s & vbTab & TFLines(i).Note
    oTF.WriteToTextFile s
Next

    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(Me.TFLines.Count)
    oTF.WriteToTextFile "TotalQtyItems" & vbTab & TotalQtyItemsF
    oTF.WriteToTextFile "ExtExVAT" & vbTab & ExtExVATF
    oTF.WriteToTextFile "DiscExVAT" & vbTab & DiscExVATF
    oTF.WriteToTextFile "ExtLessDiscExVAT" & vbTab & ExtLessDiscExVATF
    oTF.WriteToTextFile "TotalCostExVAT" & vbTab & TotalCostExVATF
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
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
    ErrorIn "a_PO.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub
Public Property Get SendersEmail() As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
        SendersEmail = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).EMail
    If strT > "" Then
        SendersEmail = strT
    Else
        SendersEmail = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF.SendersEmail"
End Property

Public Property Get StatusForPrinting() As String
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
    ErrorIn "a_TF.StatusForPrinting"
End Property

Public Property Get Signature(Optional bSkipBrackets As Boolean) As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
        Signature = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).Signature
    If bSkipBrackets Then
        If strT > "" Then
            Signature = strT
        Else
            Signature = ""
        End If
    Else
        If strT > "" Then
            Signature = " (" & strT & ")"
        Else
            Signature = ""
        End If
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.Signature"
End Property

Public Function IsOktoExport() As Integer
Dim iCnt As Integer
Dim oTFL As a_TFL

    iCnt = 0
    For Each oTFL In Me.TFLines
        If Left(oTFL.EAN, 1) = "2" Then
            iCnt = iCnt + 1
        End If
    Next
    IsOktoExport = iCnt
End Function
