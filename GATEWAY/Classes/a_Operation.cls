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
Attribute VB_Exposed = True
Option Explicit
'Public Enum OpTypes
'    enLoyaltyScheme = 11
'    enStockSharing = 12
'    enNielsen = 9
'    enExportGroup = 20
'End Enum
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
Private WithEvents objValid As z_BrokenRules
Attribute objValid.VB_VarHelpID = -1


Private lngID As Long

Public Sub BeginEdit()
    On Error GoTo errHandler
    If flgEditing Then Err.Raise 445
    LSet udtSave = udtProps
    flgEditing = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.BeginEdit", , EA_NORERAISE
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
    ErrorIn "a_Operation.CancelEdit", , EA_NORERAISE
End Sub
Public Sub ApplyEdit(Optional lngResult As Long)
    On Error GoTo errHandler
Dim iResult As Long

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
'ERR_Handler:
'    oError.SetError Err, Error, Now(), "a_Operation:ApplyEdit", "", ""
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.ApplyEdit(lngResult)", lngResult, EA_NORERAISE
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set objValid = New z_BrokenRules
    flgNew = True
    flgEditing = False
    udtProps.Result = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Class_Initialize", , EA_NORERAISE
End Sub
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (objValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsValid", , EA_NORERAISE
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set objValid = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Class_Terminate", , EA_NORERAISE
End Sub

Private Sub objValid_BrokenRule(Errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.objValid_BrokenRule(Errors)", Errors, EA_NORERAISE
    
End Sub

Private Sub objValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid(True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.objValid_NoBrokenRules", , EA_NORERAISE
    
End Sub

Public Sub Load(lngResult As Long, PID As Long)
    On Error GoTo errHandler
Dim tmperr As String
    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    Fetch lngResult, PID:=PID
    flgNew = False

EXIT_Handler:
'ERR_Handler:
'    tmperr = Err
'    If tmperr = 99999 Then
'        lngResult = tmperr
'    Else 'Will not be handled by the GUI
'        oError.SetError Err, Error, Now(), "a_Operation:Load", "", ""
'        GoTo EXIT_Handler
'    End If
'    GoTo EXIT_Handler
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Load(lngResult,PID)", Array(lngResult, PID), EA_NORERAISE
    
End Sub
Public Sub Delete()
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 445
    flgDeleted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Delete", , EA_NORERAISE
    
End Sub
Private Sub DeleteObject(lngResult As Long, ID As Long)
    On Error GoTo errHandler

  Dim strSQL As String
  Dim lngRecs As Long
  Dim tmperr As String
  
    strSQL = "DELETE * FROM tblOperation WHERE Operation_ID=" & ID
    oPC.COShort.Execute strSQL
    
EXIT_Handler:
'ERR_Handler:
'    oError.SetError Err, Error, Now(), "a_Operation:DeleteObject", "", ""
'    GoTo EXIT_Handler
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.DeleteObject(lngResult,ID)", Array(lngResult, ID)
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
    IsDeleted = flgDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsDeleted", , EA_NORERAISE
    
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
    IsNew = flgNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsNew", , EA_NORERAISE
    
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = flgDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.IsDirty", , EA_NORERAISE
    
End Property
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = udtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.ID", , EA_NORERAISE
    
End Property
Public Property Get StartedAt() As Date
    On Error GoTo errHandler
    StartedAt = udtProps.StartedAt
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.StartedAt", , EA_NORERAISE
    
End Property
Public Property Let StartedAt(Val As Date)
    On Error GoTo errHandler
    If IsDate(Val) Then udtProps.StartedAt = Val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.StartedAt(val)", Val, EA_NORERAISE
    
End Property
Public Property Get Endedat() As Date
    On Error GoTo errHandler
    Endedat = udtProps.Endedat
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Endedat", , EA_NORERAISE
    
End Property
Public Property Let Endedat(Val As Date)
    On Error GoTo errHandler
    If IsDate(Val) Then udtProps.Endedat = Val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Endedat(val)", Val, EA_NORERAISE
    
End Property
Public Property Get NominalDate() As Date
    On Error GoTo errHandler
    NominalDate = udtProps.NominalDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.NominalDate", , EA_NORERAISE
    
End Property
Public Property Let NominalDate(Val As Date)
    On Error GoTo errHandler
    If IsDate(Val) Then udtProps.NominalDate = Val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.NominalDate(val)", Val, EA_NORERAISE
    
End Property
Public Property Get TypeName() As String
    On Error GoTo errHandler
    TypeName = Trim(udtProps.TypeName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.TypeName", , EA_NORERAISE
    
End Property
Public Property Get Fullreport() As String
    On Error GoTo errHandler
    Fullreport = Trim(udtProps.Fullreport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Fullreport", , EA_NORERAISE
    
End Property
Public Property Let Fullreport(Val As String)
    On Error GoTo errHandler
    udtProps.Fullreport = Trim(Val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Fullreport(val)", Val, EA_NORERAISE
    
End Property
Public Property Let TypeID(Val As OpTypes)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    udtProps.Type = Val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.TypeID(val)", Val, EA_NORERAISE
    
End Property
Public Property Get TypeDesc() As String
    On Error GoTo errHandler
    Select Case udtProps.Type
        Case NielsenSales: TypeDesc = "Update sales data to Nielsen"
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.TypeDesc", , EA_NORERAISE
    
End Property

Public Property Get OperatorID() As Long
    On Error GoTo errHandler
    OperatorID = udtProps.OperatorID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.OperatorID", , EA_NORERAISE
    
End Property
Public Property Let OperatorID(Val As Long)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    objValid.RuleBroken "Operator", (Val = 0)
    udtProps.OperatorID = Val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.OperatorID(val)", Val, EA_NORERAISE
    
End Property
Public Property Get Result() As Long
    On Error GoTo errHandler
    Result = udtProps.Result
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Result", , EA_NORERAISE
    
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
    ErrorIn "a_Operation.ResultF", , EA_NORERAISE
    
End Property
Public Property Let Result(Val As Long)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    udtProps.Result = Val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Result(val)", Val, EA_NORERAISE
    
End Property
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property
Private Sub Save(lngResult As Long)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strNewCode
    
    Set rs = New ADODB.Recordset
    If flgNew Then
        rs.Open "Select * from tOperation", oPC.COShort, adOpenDynamic, adLockOptimistic
        rs.AddNew
    Else
        rs.Open "Select * From tOperation", oPC.COShort, adOpenDynamic, adLockOptimistic
        rs.Find "[OP_ID] = " & udtProps.ID
    End If
    With rs
        .Fields("OP_EndedAT") = FND(udtProps.Endedat)
        .Fields("OP_StartedAt") = FND(udtProps.StartedAt)
        .Fields("OP_NominalDate") = FND(udtProps.NominalDate)
        .Fields("OP_Result") = FNN(udtProps.Result)
        .Fields("OP_Type") = FNS(udtProps.Type)
        .Fields("OP_FullReport") = FNS(udtProps.Fullreport)
        .Update
        If flgNew Then
            .MoveLast
            udtProps.ID = .Fields("OP_ID")
        End If
        .Close
    End With
    
EXIT_Handler:
    Set rs = Nothing
'ERR_Handler:
'        oError.SetError Err, Error, Now(), "a_Operation:Save", "", ""
'        GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Operation.Save(lngResult)", lngResult
End Sub
Private Sub Fetch(lngResult As Long, PID As Long)
    On Error GoTo errHandler

Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tOPERATION", oPC.COShort, adOpenDynamic, adLockReadOnly
        
    rs.Find "[OP_ID] = " & PID
    If rs.EOF And rs.BOF Then
        lngResult = 1
        GoTo EXIT_Handler
    Else
        lngResult = 0
    End If
    With rs
        udtProps.ID = .Fields("OP_ID")
        udtProps.Type = FNS(.Fields("OP_Type"))
        udtProps.Endedat = FND(.Fields("OP_EndedAt"))
        udtProps.StartedAt = FND(.Fields("OP_StartedAt"))
        udtProps.NominalDate = FND(.Fields("OP_NominalDate"))
        udtProps.Result = FNN(.Fields("OP_Result"))
        udtProps.OperatorID = FNN(.Fields("OP_StartedByID"))
    '    udtProps.ResultName = FNN(.Fields("ResultName"))
        udtProps.Fullreport = FNS(.Fields("OP_FullReport"))
        .Close
    End With
EXIT_Handler:
    Set rs = Nothing
'ERR_Handler:
'    oError.SetError Err, Error, Now(), "a_Operation:Fetch", "", ""
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
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
    ErrorIn "a_Operation.LoadByNominalDate(pEnterDate)", pEnterDate, EA_NORERAISE
    
End Sub
