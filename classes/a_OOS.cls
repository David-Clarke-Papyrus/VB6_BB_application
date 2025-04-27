VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_OOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type Props
    ID As Long
    Code As String * 10
    WeekNo As Integer
    Date As Date
    OperatorID As Long
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
Public objDetails As ch_OOSLine

Private tlOperators As z_TextList

Private lngID As Long

Public Sub BeginEdit()
    On Error GoTo errHandler
    If flgEditing Then Err.Raise 445
    LSet udtSave = udtProps
    objDetails.BeginEdit
    flgEditing = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.BeginEdit"
End Sub
Public Sub CancelEdit()
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 445
    flgEditing = False
    flgDeleted = False
    LSet udtProps = udtSave
    objDetails.CancelEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.CancelEdit"
End Sub
Public Sub ApplyEdit(lngResult As Long)
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
    objDetails.ApplyEdit udtProps.ID
    flgDirty = False
    flgEditing = False
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.ApplyEdit(lngResult)", lngResult
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set objValid = New z_BrokenRules
    flgNew = True
    flgEditing = False

    Set tlOperators = New z_TextList
    tlOperators.Load "Operators_qsel"

    Set objDetails = New ch_OOSLine

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.Class_Initialize"
End Sub
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (objValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.IsValid"
End Property

Private Sub objValid_BrokenRule(pErrors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.objValid_BrokenRule(pErrors)", pErrors
End Sub

Private Sub objValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid(True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.objValid_NoBrokenRules"
End Sub

Public Sub Load(lngResult As Long, Optional PID As Long, Optional pCODE As String)
    On Error GoTo errHandler
    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    If PID > 0 Then
        Fetch lngResult, PID:=PID
    ElseIf Len(pCODE) > 0 Then
        Fetch lngResult, pCODE:=pCODE
    End If
    objDetails.Load udtProps.ID
    flgNew = False

EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.Load(lngResult,PID,pCODE)", Array(lngResult, PID, pCODE)
End Sub
Public Sub Delete()
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 445
    flgDeleted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.Delete"
End Sub
Private Sub DeleteObject(lngResult As Long, ID As Long)
    On Error GoTo errHandler

  Dim strSQL As String
  Dim lngRecs As Long
  
    strSQL = "DELETE * FROM tblOOS WHERE OOS_ID=" & ID
    oPC.CO.Execute strSQL
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.DeleteObject(lngResult,ID)", Array(lngResult, ID)
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
    IsDeleted = flgDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
    IsNew = flgNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.IsNew"
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = flgDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.IsDirty"
End Property
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = udtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.ID"
End Property
Public Property Get Code() As String
    On Error GoTo errHandler
    Code = Trim$(udtProps.Code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.Code"
End Property
Public Property Get WeekNo() As Integer
    On Error GoTo errHandler
    WeekNo = udtProps.WeekNo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.WeekNo"
End Property
Public Property Let WeekNo(val As Integer)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    udtProps.WeekNo = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.WeekNo(val)", val
End Property
Public Property Get OperatorID() As Long
    On Error GoTo errHandler
    OperatorID = udtProps.OperatorID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.OperatorID"
End Property
Public Property Let OperatorID(val As Long)
    On Error GoTo errHandler
    If Not flgEditing Then Err.Raise 383
    objValid.RuleBroken "Operator", (val = 0)
    udtProps.OperatorID = val
    flgDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.OperatorID(val)", val
End Property
Public Property Get Operators() As z_TextList
    Set Operators = tlOperators
End Property
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property
Private Sub Save(lngResult As Long)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim objCode As New z_Code
Dim strNewCode

    If flgNew Then
        rs.Open "OOS_qsel", oPC.CO, adOpenDynamic, adLockOptimistic
        rs.AddNew
        udtProps.Code = "OS" & objCode.GetNextCode("OOS", "GetandWrite")
    Else
        rs.Open "OOS_qsel", oPC.CO, adOpenDynamic, adLockOptimistic
        rs.Find "[OOS_ID] = " & udtProps.ID
    End If
    With rs
        If HasData(udtProps.Code) Then .Fields("OOS_Code") = Trim$(udtProps.Code)
        If HasData(udtProps.Date) Then .Fields("OOS_Date") = Trim$(udtProps.Date)
        If HasData(udtProps.WeekNo) Then .Fields("OOS_Weekno") = Trim$(udtProps.WeekNo)
        .Update
        If flgNew Then
            .MoveLast
            udtProps.ID = .Fields("OOS_ID")
        End If
        .Close
    End With
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.Save(lngResult)", lngResult
End Sub
Private Sub Fetch(lngResult As Long, Optional PID As Long, Optional pCODE, Optional pTRID As Long)
    On Error GoTo errHandler

Dim rs As ADODB.Recordset

    rs.Open "OOS_qsel", oPC.CO, adOpenDynamic, adLockReadOnly
        
    If PID > 0 Then
        rs.Find "[CashSale_ID] = " & PID
    ElseIf Not IsMissing(pCODE) Then
        rs.Find "[TR_Code] = '" & pCODE & "'"
    Else
        lngResult = 99
        GoTo EXIT_Handler
    End If
    If rs.EOF And rs.BOF Then
        lngResult = 1
        GoTo EXIT_Handler
    Else
        lngResult = 0
    End If
    With rs
        udtProps.ID = .Fields("OOS_ID")
        udtProps.Code = FNS(.Fields("OOS_Code"))
        udtProps.Date = .Fields("OOS_Date")
        udtProps.OperatorID = FNN(.Fields("OOS_OperatorID"))
        .Close
    End With
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOS.Fetch(lngResult,PID,pCODE,pTRID)", Array(lngResult, PID, pCODE, pTRID)
End Sub
