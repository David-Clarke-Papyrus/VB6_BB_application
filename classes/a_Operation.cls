VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Operation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Props
    ID As Long
    Type As OpTypes
    StartedAt As Date
    Endedat As Date
    NominalDate As Date
    Result As Long
    OperatorID As Long
    TypeName As String * 15
    ResultName As String * 15
    Fullreport As String * 2000
    
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Dim udtProps As Props
Dim udtSave As Props
Event Valid(IsValid As Boolean)

Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private tlOperators As z_TextList

Private lngID As Long

Public Sub BeginEdit()
    On Error GoTo errHandler
    If flgEditing Then Err.Raise 445
    LSet udtSave = udtProps
    flgEditing = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.BeginEdit"
End Sub
Public Sub CancelEdit()
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 445
    flgEditing = False
    flgDeleted = False
    LSet udtProps = udtSave
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.CancelEdit"
End Sub
Public Sub ApplyEdit(Optional lngResult As Long)
    On Error GoTo errHandler
Dim iresult As Long

    If Not flgEditing Then Err.Raise 445
    If flgDeleted Then
        'code to delete objects data goes here
   '     DeleteObject udtProps.ID, iresult
        flgNew = True
        flgDeleted = False
    ElseIf flgDirty Or flgNew Then
        If Not IsValid Then Err.Raise 445
        'save object to database if appropriate
        Save lngResult
        If lngResult <> 0 Then
            GoTo EXIT_Handler
        End If
        LSet udtSave = udtProps 'Sets the save properties to what has been saved - checkpoint
        flgNew = False
    End If
    flgDirty = False
    flgEditing = False
    
EXIT_Handler:
'Err_Handler:
'    MsgBox "a_OPeration Applyedit " & or
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.ApplyEdit(lngResult)", lngResult
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mobjValid = New z_BrokenRules
    Set tlOperators = New z_TextList
    flgNew = True
    flgEditing = False
    udtProps.Result = 1
    tlOperators.Load ltStaff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Class_Initialize"
End Sub
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set mobjValid = Nothing
    Set tlOperators = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.mobjValid_BrokenRule(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid(True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.mobjValid_NoBrokenRules"
End Sub

Public Sub Load(lngResult As Long, PID As Long)
    On Error GoTo errHandler
Dim tmp As String
    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    Fetch lngResult, PID:=PID
    flgNew = False

EXIT_Handler:
'Err_Handler:
'    tmp =
'    If tmp = 99999 Then
'        lngResult = tmp
'    Else 'Will not be handled by the GUI
'        MsgBox "a_OPeration Load " & or
'    End If
'    GoTo EXIT_Handler
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Load(lngResult,PID)", Array(lngResult, PID)
End Sub
Public Sub Delete()
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 445
    flgDeleted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Delete"
End Sub
Private Sub DeleteObject(lngResult As Long, ID As Long)
    On Error GoTo errHandler

  Dim strSQL As String
  Dim lngRecs As Long
  Dim tmp As String
  
    strSQL = "DELETE * FROM tblOperation WHERE Operation_ID=" & ID
    oPC.COShort.execute strSQL
    
EXIT_Handler:
'Err_Handler:
'    If (( = DB_EXISTS) Or ( = DB_MISSING) Or ( = DB_REF) Or ( = DB_FIELDCONSTRAINT)) Then
'        lngResult =
'    Else
'        MsgBox "a_Operation: Deleteobject " & or
'    End If
'    GoTo EXIT_Handler
'
    Exit Sub
errHandler:
    ErrPreserve
    If ((Err = DB_EXISTS) Or (Err = DB_MISSING) Or (Err = DB_REF) Or (Err = DB_FIELDCONSTRAINT)) Then
        lngResult = Err
        Resume EXIT_Handler
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.DeleteObject(lngResult,ID)", Array(lngResult, ID)
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
    IsDeleted = flgDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
    IsNew = flgNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsNew"
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = flgDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsDirty"
End Property
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = udtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.ID"
End Property
Public Property Get StartedAt() As Date
    On Error GoTo errHandler
    StartedAt = udtProps.StartedAt
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.StartedAt"
End Property
Public Property Let StartedAt(val As Date)
    On Error GoTo errHandler
    If IsDate(val) Then udtProps.StartedAt = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.StartedAt(val)", val
End Property
Public Property Get Endedat() As Date
    On Error GoTo errHandler
    Endedat = udtProps.Endedat
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Endedat"
End Property
Public Property Let Endedat(val As Date)
    On Error GoTo errHandler
    If IsDate(val) Then udtProps.Endedat = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Endedat(val)", val
End Property
Public Property Get NominalDate() As Date
    On Error GoTo errHandler
    NominalDate = udtProps.NominalDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.NominalDate"
End Property
Public Property Let NominalDate(val As Date)
    On Error GoTo errHandler
    If IsDate(val) Then udtProps.NominalDate = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.NominalDate(val)", val
End Property
Public Property Get TypeName() As String
    On Error GoTo errHandler
    TypeName = Trim(udtProps.TypeName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.TypeName"
End Property
Public Property Get Fullreport() As String
    On Error GoTo errHandler
    Fullreport = Trim(udtProps.Fullreport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Fullreport"
End Property
Public Property Let Fullreport(val As String)
    On Error GoTo errHandler
    udtProps.Fullreport = Trim(val)
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Fullreport(val)", val
End Property
Public Property Let TypeID(val As OpTypes)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    udtProps.Type = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.TypeID(val)", val
End Property
Public Property Get TypeDesc() As String
    On Error GoTo errHandler
    Select Case udtProps.Type
        Case Dayend: TypeDesc = "Daily update"
        Case Compact: TypeDesc = "Compact"
        Case ClearOrders: TypeDesc = "Clear orders"
        Case Export: TypeDesc = "Export"
        Case SupplierChange: TypeDesc = "Supplier change"
        Case SimplifyTitles: TypeDesc = "Remove 'The' etc."
        Case AllocationStockToOrders: TypeDesc = "Order fulfilment"
        Case NielsenSales: TypeDesc = "Export sales data"
        Case ClearOldData: TypeDesc = "Clear old data"
        Case BookDataWash: TypeDesc = "Book data wash"
        Case LoyaltyScheme: TypeDesc = "Export loyalty customer"
        Case StockSharing: TypeDesc = "Export stock-sharing"
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.TypeDesc"
End Property
'''''    Dayend = 1
'''''    Compact = 2
'''''    ClearOrders = 3
'''''    Export = 4
'''''    SupplierChange = 5
'''''    SimplifyTitles = 6
'''''    ClearOldData = 7
'''''    AllocationStockToOrders = 8
'''''    NielsenSales = 9
'''''    BookDataWash = 10
'''''    LoyaltyScheme = 11
'''''    StockSharing = 12

Public Property Get OperatorID() As Long
    On Error GoTo errHandler
    OperatorID = udtProps.OperatorID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.OperatorID"
End Property
Public Property Let OperatorID(val As Long)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    mobjValid.RuleBroken "Operator", (val = 0)
    udtProps.OperatorID = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.OperatorID(val)", val
End Property
Public Property Get Result() As Long
    On Error GoTo errHandler
    Result = udtProps.Result
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Result"
End Property
Public Property Get ResultF() As String
    On Error GoTo errHandler
    Select Case udtProps.Result
    Case 1
        ResultF = "OK"
    Case 2
        ResultF = "Failed"
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.ResultF"
End Property
Public Property Let Result(val As Long)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    udtProps.Result = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Result(val)", val
End Property
Public Property Get Operators() As z_TextList
    Set Operators = tlOperators
End Property
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property
Private Sub Save(lngResult As Long)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim objCode As z_Code
Dim strNewCode
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    Set objCode = New z_Code
    If flgNew Then
        rs.open "Select * from tOperation", oPC.COShort, adOpenDynamic, adLockOptimistic
        rs.AddNew
    Else
        rs.open "Select * From tOperation", oPC.COShort, adOpenDynamic, adLockOptimistic
        rs.Find "[OP_ID] = " & udtProps.ID
    End If
    With rs
        .fields("OP_EndedAT") = FND(udtProps.Endedat)
        .fields("OP_StartedAt") = FND(udtProps.StartedAt)
        .fields("OP_NominalDate") = FND(udtProps.NominalDate)
        .fields("OP_Result") = FNN(udtProps.Result)
        .fields("OP_Type") = FNN(udtProps.Type)
        .fields("OP_StartedByID") = FNN(udtProps.OperatorID)
        .fields("OP_FullReport") = FNS(udtProps.Fullreport)
        .Update
        If flgNew Then
            .MoveLast
            udtProps.ID = .fields("OP_ID")
        End If
        .Close
    End With

EXIT_Handler:
    Set rs = Nothing
    Set objCode = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If ((Err = DB_EXISTS) Or (Err = DB_MISSING) Or (Err = DB_REF) Or (Err = DB_FIELDCONSTRAINT)) Then
        lngResult = Err
        Resume EXIT_Handler
    End If
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Save(lngResult)", lngResult
End Sub
Private Sub Fetch(lngResult As Long, PID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.open "SELECT * FROM tOPERATION", oPC.COShort, adOpenDynamic, adLockReadOnly
        
    rs.Find "[OP_ID] = " & PID
    If rs.eof And rs.BOF Then
        lngResult = 1
        GoTo EXIT_Handler
    Else
        lngResult = 0
    End If
    With rs
        udtProps.ID = .fields("OP_ID")
        udtProps.Type = FNS(.fields("OP_Type"))
        udtProps.Endedat = FND(.fields("OP_EndedAt"))
        udtProps.StartedAt = FND(.fields("OP_StartedAt"))
        udtProps.NominalDate = FND(.fields("OP_NominalDate"))
        udtProps.Result = FNN(.fields("OP_Result"))
        udtProps.OperatorID = FNN(.fields("OP_StartedByID"))
    '    udtProps.ResultName = FNN(.Fields("ResultName"))
        udtProps.Fullreport = FNS(.fields("OP_FullReport"))
        .Close
    End With
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Set rs = Nothing

    Exit Sub
errHandler:
    
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Fetch(lngResult,PID)", Array(lngResult, PID)
End Sub

Public Sub LoadByNominalDate(ByRef pEnterDate As Date)
    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oBatch As z_SQL
'Dim retval
'    Set rs = New ADODB.Recordset
'    Set oBatch = New z_SQL
'
'    retval = oBatch.RunProc("q_GetNominalDate_1", Array(), "")
'    retval = oBatch.RunGetRecordset("q_GetNominalDate_2", Query, Array(pEnterDate), "", rs)
'    If Not rs.EOF Then
'        pEnterDate = rs.Fields("O_NominalDate")
'    End If
'
'    Set oBatch = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.LoadByNominalDate(pEnterDate)", pEnterDate
End Sub
