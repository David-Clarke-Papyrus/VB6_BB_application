VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Prodset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum ProdSetTypes
    DeletionsfromABE = 1
    AdditionstoABE = 2
    ReplacementtoABE = 3
    General = 3
End Enum
Private Type Props
    ID As Long
    TypeID As ProdSetTypes
    CreatedDate As Date
    OperatorID As Long
    Description As String * 120
    
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

Private tlOperators As z_TextList

Private lngID As Long

Public Sub BeginEdit()
    If flgEditing Then Err.Raise 445
    LSet udtSave = udtProps
    flgEditing = True
End Sub
Public Sub CancelEdit()
    If Not flgEditing Then Err.Raise 445
    flgEditing = False
    flgDeleted = False
    LSet udtProps = udtSave
End Sub
Public Sub ApplyEdit(lngResult As Long)
On Error GoTo ERR_Handler
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
    Exit Sub
ERR_Handler:
    tmpErr = Err
    tmpError = Error
    oError.SetError tmpErr, tmpError, Now(), "a_Prodset", "ApplyEdit", "unknown"
    Err.Raise tmpErr
    GoTo EXIT_Handler
End Sub

Private Sub Class_Initialize()
    Set objValid = New z_BrokenRules
    udtProps.CreatedDate = Now()
    flgNew = True
    flgEditing = False
    Set tlOperators = New z_TextList
    tlOperators.Load "Operators_qsel"

End Sub
Public Property Get IsValid() As Boolean
    IsValid = (objValid.Count = 0)
End Property

Private Sub objValid_BrokenRule(errors As String)
    RaiseEvent Valid(False)
End Sub

Private Sub objValid_NoBrokenRules()
    RaiseEvent Valid(True)
End Sub

Public Sub Load(lngResult As Long, PID As Long)
On Error GoTo ERR_Handler
    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    Fetch lngResult, PID:=PID
    flgNew = False

EXIT_Handler:
    Exit Sub
ERR_Handler:
    tmpErr = Err
    tmpError = Error
    If tmpErr = 99999 Then
        lngResult = tmpErr
    Else 'Will not be handled by the GUI
        lngResult = 0
        oError.SetError tmpErr, tmpError, Now(), "a_Prodset", "Load", ""
    End If
    GoTo EXIT_Handler

End Sub
Public Sub Delete(PID As Long)
Dim lngResult As Long
    DeleteObject lngResult, PID
    flgDeleted = True
End Sub
Private Sub DeleteObject(lngResult As Long, PID As Long)
On Error GoTo ERR_Handler

  Dim strSQL As String
  Dim lngRecs As Long
    strSQL = "DELETE * FROM Product_ProdSet WHERE P_PS_PS_ID = " & PID
    oPC.CO.Execute strSQL
    strSQL = "DELETE * FROM Prodset WHERE Prodset_ID=" & PID
    oPC.CO.Execute strSQL
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    If ((Err = DB_EXISTS) Or (Err = DB_MISSING) Or (Err = DB_REF) Or (Err = DB_FIELDCONSTRAINT)) Then
        lngResult = Err
    Else
        tmpErr = Err
        tmpError = Error
        oError.SetError tmpErr, tmpError, Now(), "a_Prodset", "DeleteObject", "unknown"
        Err.Raise tmpErr
    End If
    GoTo EXIT_Handler

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
Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Public Property Get CreatedDate() As Date
    CreatedDate = udtProps.CreatedDate
End Property
Public Property Let CreatedDate(Val As Date)
    If Not flgEditing Then Err.Raise 383
    If IsDate(Val) Then udtProps.CreatedDate = Val
    flgDirty = True
End Property
Public Property Get Description() As String
    Description = Trim(udtProps.Description)
End Property
Public Property Let Description(Val As String)
    If Not flgEditing Then Err.Raise 383
    udtProps.Description = Trim(Val)
    flgDirty = True
End Property
Public Property Get OperatorID() As Long
    OperatorID = udtProps.OperatorID
End Property
Public Property Let OperatorID(Val As Long)
    If Not flgEditing Then Err.Raise 383
    objValid.RuleBroken "Operator", (Val = 0)
    udtProps.OperatorID = Val
    flgDirty = True
End Property
Public Property Get TypeID() As Long
    TypeID = udtProps.TypeID
End Property
Public Property Get TypeAsText() As String
    Select Case udtProps.TypeID
        Case DeletionsfromABE: TypeDescription = "Deletions from ABE"
        Case AdditionstoABE: TypeDescription = "Additions to ABE"
        Case ReplacementtoABE: TypeDescription = "Replacement to ABE"
        Case General: TypeDescription = "General"
    End Select
End Property
Public Property Let TypeDescription(Val As ProdSetTypes)
    If Not flgEditing Then Err.Raise 383
    udtProps.TypeID = Val
    flgDirty = True
End Property

Public Property Let TypeID(Val As Long)
    If Not flgEditing Then Err.Raise 383
    udtProps.TypeID = Val
    flgDirty = True
End Property
Public Property Get Operators() As z_TextList
    Set Operators = tlOperators
End Property
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property
Private Sub Save(lngResult As Long)
On Error GoTo ERR_Handler
Dim rs As New ADODB.Recordset
Dim objCode As New z_Code
Dim strNewCode

    If flgNew Then
        rs.Open "Select * from Prodset", oPC.CO, adOpenDynamic, adLockOptimistic
        rs.AddNew
    Else
        rs.Open "Select * From Prodset", oPC.CO, adOpenDynamic, adLockOptimistic
        rs.Find "[Operation_ID] = " & udtProps.ID
    End If
    With rs
        If HasData(udtProps.CreatedDate) Then .Fields("PS_CreatedDate") = Trim$(udtProps.CreatedDate)
     '   If HasData(udtProps.OperatorID) Then .Fields("O_StartedBy") = Trim$(udtProps.OperatorID)
        If HasData(udtProps.Description) Then .Fields("PS_Description") = Trim$(udtProps.Description)
        .Update
        If flgNew Then
            .MoveLast
            udtProps.ID = .Fields("Prodset_ID")
        End If
        .Close
    End With
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    If ((Err = DB_EXISTS) Or (Err = DB_MISSING) Or (Err = DB_REF) Or (Err = DB_FIELDCONSTRAINT)) Then
        lngResult = Err
    Else
        oError.SetError Err, Error, Now(), "a_Prodset", "Save", "unknown"
        Err.Raise Err
    End If
End Sub
Private Sub Fetch(lngResult As Long, PID As Long)
On Error GoTo ERR_Handler

Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "GetProdset", oPC.CO, adOpenDynamic, adLockReadOnly
        
    rs.Find "[Prodset_ID] = " & PID
    If rs.EOF And rs.BOF Then
        lngResult = 1
        GoTo EXIT_Handler
    Else
        lngResult = 0
    End If
    With rs
        udtProps.ID = .Fields("Prodset_ID")
        udtProps.CreatedDate = FixNullsDate(.Fields("PS_CreatedDate"))
        udtProps.OperatorID = FixNullsString(.Fields("O_StartedBy"))
        If HasData(.Fields("PS_Description")) Then udtProps.Description = .Fields("PS_Description")
        .Close
    End With
EXIT_Handler:
    Exit Sub
ERR_Handler:
    tmpErr = Err
    tmpError = Error
    oError.SetError tmpErr, tmpError, Now(), "a_Prodset", "Fetch", "unknown"
    Err.Raise tmpErr
    GoTo EXIT_Handler
End Sub
