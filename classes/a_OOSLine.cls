VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_OOSLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type Props
    ID As Long
    ProductID As Long
    InOut As Integer
    CheckedQty As Integer
    OOSID As Long
    
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Dim udtProps As Props
Dim udtSave As Props
Dim udtChild As Props
Event Valid(IsValid As Boolean)

Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean
Private flgChild As Boolean
Private flgChildEditing As Boolean
Private objParent As ch_OOSLine
Private WithEvents objValid As z_BrokenRules
Attribute objValid.VB_VarHelpID = -1
Public Sub BeginEdit()
    If flgChild Then If Not flgChildEditing Then Err.Raise 445
    If flgEditing Then Err.Raise 445
    LSet udtSave = udtProps
    flgEditing = True
End Sub
Public Sub CancelEdit()
    If Not flgEditing Then Err.Raise 445
    flgEditing = False
    flgDeleted = False
    'Restor object state
    LSet udtProps = udtSave
    flgDirty = False
End Sub
Public Sub ApplyEdit(lngResult As Long)
    On Error GoTo errHandler
Dim flgSuccess As Boolean

  If flgChild Then If Not flgChildEditing Then Err.Raise 445
  If Not flgChild And flgNew Then Err.Raise 445
  If Not flgEditing Then Err.Raise 445
  lngResult = 0
  If Not flgChild Then
    If flgDeleted Then
        'code to delete objects data goes here
        lngResult = 0
        DeleteObject lngResult, udtProps.ID  'error here error here
        If lngResult <> 0 Then
            GoTo EXIT_Handler
        End If
        flgNew = True
        flgDeleted = False
    ElseIf flgDirty Or flgNew Then
        If Not IsValid Then Err.Raise 445
        lngResult = 0
        Save lngResult
        If lngResult <> 0 Then
            GoTo EXIT_Handler
        End If
        LSet udtSave = udtProps
        flgNew = False
    End If
  flgDirty = False
  Else
    If flgNew Then
        If Not objParent Is Nothing Then
            flgSuccess = objParent.AddOOSL(Me)
            If Not flgSuccess Then lngResult = 50
            Set objParent = Nothing
        End If
    End If
  End If
  flgEditing = False
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOSLine.ApplyEdit(lngResult)", lngResult
End Sub

Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Public Property Let CheckedQty(val As Integer)
    udtProps.CheckedQty = val
    flgDirty = True
End Property
Public Property Get CheckedQty() As Integer
    CheckedQty = udtProps.CheckedQty
End Property
Public Property Let InOut(val As Integer)
    udtProps.InOut = val
    flgDirty = True
End Property
Public Property Get InOut() As Integer
    InOut = udtProps.InOut
End Property
Public Property Let OOSID(val As Long)
    udtProps.OOSID = val
    flgDirty = True
End Property
Public Property Get OOSID() As Long
    OOSID = udtProps.OOSID
End Property
Public Property Let ProductID(val As Long)
    udtProps.ProductID = val
    flgDirty = True
End Property
Public Property Get ProductID() As Long
    ProductID = udtProps.ProductID
End Property

Private Sub objValid_BrokenRule(pErrors As String)
    RaiseEvent Valid(False)
End Sub

Private Sub objValid_NoBrokenRules()
    RaiseEvent Valid(True)
End Sub
Public Sub Save(lngResult As Long)
    On Error GoTo errHandler

  Dim rs As New ADODB.Recordset
  Dim strSQL As String
  
  If flgNew Then
    strSQL = "tblOOSLine"
    rs.Open strSQL, oPC.CO, adOpenDynamic, adLockOptimistic
    rs.AddNew
  Else
    strSQL = "SELECT * FROM tblOOSLine WHERE OOSL_ID = " & udtProps.ID
    rs.Open strSQL, oPC.CO, adOpenDynamic, adLockOptimistic
  End If
  With rs
    .Fields("OOSL_P_ID") = udtProps.ProductID
    .Fields("OOSL_OOS_ID") = udtProps.OOSID
    .Fields("OOSL_Inout") = udtProps.InOut
    .Fields("OOSL_CheckedQty") = udtProps.CheckedQty
    .Update
    If flgNew Then
        .MoveLast
        udtProps.ID = .Fields("OOSL_ID")
    End If
    .Close
  End With
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOSLine.Save(lngResult)", lngResult
End Sub
Private Sub Class_Initialize()
    Set objValid = New z_BrokenRules
    flgNew = True
    
End Sub

Public Property Get IsValid() As Boolean
    IsValid = (objValid.Count = 0)
End Property
Public Sub Load(OOSLID As Long)
    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    
    Fetch OOSLID
    flgNew = False
End Sub
Public Sub Delete()
    If Not flgEditing Then Err.Raise 445
    flgDeleted = True
End Sub
Public Sub Undelete()
    flgDeleted = False
End Sub
Public Property Get IsDeleted() As Boolean
    IsDeleted = flgDeleted
End Property
Public Property Get IsNew() As Boolean
    IsNew = flgNew
End Property
Public Property Get IsDirty() As Boolean
    IsDirty = flgDirty
End Property
Friend Sub SetAsChild(obj As ch_OOSLine)
    If flgNew Then Set objParent = obj
    flgChild = True
End Sub

Friend Sub ChildBeginEdit()
  If flgChildEditing Then Err.Raise 445
  
  LSet udtChild = udtProps
  flgChildEditing = True
End Sub

Friend Sub ChildCancelEdit()
  If Not flgChildEditing Then Err.Raise 445
  
  flgChildEditing = False
  flgDeleted = False
  ' restore object state
  LSet udtProps = udtChild
End Sub

Friend Sub ChildApplyEdit(ID As Long)
Dim lngResult As Long
  If Not flgChildEditing Then Err.Raise 445
  If Not IsValid Then Err.Raise 445

  If flgDeleted Then
    ' code to delete the object's data goes here
    DeleteObject lngResult, udtProps.ID
    flgNew = True
    flgDeleted = False
  ElseIf flgDirty Or flgNew Then
    ' save object to database if appropriate
    udtProps.OOSID = ID
    Save lngResult
    LSet udtChild = udtProps
    flgNew = False
  End If
  flgDirty = False
  flgChildEditing = False
End Sub

Friend Sub ChildLoad(ID As Long)
    If flgChildEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Fetch ID

  flgNew = False
  End Sub
Private Sub DeleteObject(lngResult As Long, ID As Long)
    On Error GoTo errHandler

  Dim rsLeg As New ADODB.Recordset
  Dim strSQL As String
  Dim lngRecs As Long
  
    strSQL = "DELETE * FROM tblOOSL WHERE OOSL_ID=" & ID
    oPC.CO.Execute strSQL, lngRecs, adCmdText
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOSLine.DeleteObject(lngResult,ID)", Array(lngResult, ID)
End Sub
Private Sub Fetch(pOOSLID As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim Param As ADODB.Parameter
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    Set Param = New ADODB.Parameter
    cmd.ActiveConnection = oPC.CO
    cmd.CommandText = "GetOOSLine"
    cmd.CommandType = adCmdStoredProc
    Set Param = cmd.CreateParameter("[EnterOOSLID]", adInteger, adParamInput, , pOOSLID)
    cmd.Parameters.Append Param
    Set rs = New ADODB.Recordset
    rs.Open cmd, , adOpenKeyset, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        udtProps.ID = rs.Fields("OOSL_ID")
        udtProps.CheckedQty = rs.Fields("OOSL_CheckedQty")
        udtProps.InOut = rs.Fields("OOSL_Inout")
        udtProps.ProductID = rs.Fields("OOSL_P_ID")
    End If
    rs.Close
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_OOSLine.Fetch(pOOSLID)", pOOSLID
End Sub
