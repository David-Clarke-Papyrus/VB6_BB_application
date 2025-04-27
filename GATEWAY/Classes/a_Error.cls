VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Error"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type ErrorProps
    ID As Long
    Description As String * 250
    Number As String * 100
    DateOfError As Date
    UserName As String * 100
    FormName As String * 200
    ReportName As String * 200
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Dim udtProps As ErrorProps
Dim udtSave As ErrorProps
Dim bDiagnosticsOn As Boolean
Event Valid(IsValid As Boolean)
Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean
'Private WithEvents objValid As z_BrokenRules

Public Sub BeginEdit()
  '  If flgEditing Then Err.Raise 445
    'Save object State
    LSet udtSave = udtProps
    flgEditing = True
End Sub

Public Sub CancelEdit()
 '   If Not flgEditing Then Err.Raise 445
    flgEditing = False
    flgDeleted = False
    'Restor object state
    LSet udtProps = udtSave

End Sub

Public Sub ApplyEdit()
On Error GoTo ERR_Handler
'    If Not flgEditing Then Err.Raise 445
    If flgDeleted Then
        'code to delete objects data goes here
        DeleteObject udtProps.ID
        flgNew = True
        flgDeleted = False
    ElseIf flgDirty Or flgNew Then
'        If Not IsValid Then Err.Raise 445
        Save
        LSet udtSave = udtProps
        flgNew = False
    End If
  flgDirty = False
  flgEditing = False
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
End Sub
Private Sub Class_Initialize()
    Set oError = Me
 '   Set objValid = New z_BrokenRules
    flgNew = True
    
 '   objValid.RuleBroken "Description", True
 '   objValid.RuleBroken "Number", True
End Sub

'Public Property Get IsValid() As Boolean
'    IsValid = (objValid.Count = 0)
'End Property
Public Sub SetDiagnostics(Val As Boolean)
    bDiagnosticsOn = Val
End Sub
Public Property Get DiagnosticsOn() As Boolean
    DiagnosticsOn = bDiagnosticsOn
End Property

Private Sub Class_Terminate()
    Set objValid = Nothing
End Sub

'Private Sub objValid_BrokenRule(errors As String)
'    RaiseEvent Valid(False)
'End Sub
'
'Private Sub objValid_NoBrokenRules()
'    RaiseEvent Valid(True)
'End Sub

Public Sub Load(ID As Long)
'    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    
    Fetch ID
'    objValid.RuleBroken "Description", False
'    objValid.RuleBroken "Number", False
    flgNew = False
End Sub
Public Sub Delete()
'    If Not flgEditing Then Err.Raise 445
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
Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Public Property Let Description(Val As String)
'    If Not flgEditing Then Err.Raise 383
    
'    objValid.RuleBroken "Description", (Len(Trim$(val)) = 0)
    udtProps.Description = Val
    flgDirty = True
End Property
Public Property Get Description() As String
    Description = Trim$(udtProps.Description)
End Property
Public Property Get FullDescription() As String
Dim str As String
    FullDescription = Me.FormName & vbCrLf & Me.DateOfError & vbCrLf & Me.Description
End Property
Public Property Let FormName(Val As String)
'    If Not flgEditing Then Err.Raise 383
    
    udtProps.FormName = Val
    flgDirty = True
End Property
Public Property Get FormName() As String
    FormName = Trim$(udtProps.FormName)
End Property
Public Property Let ReportName(Val As String)
'    If Not flgEditing Then Err.Raise 383
    
    udtProps.ReportName = Val
    flgDirty = True
End Property
Public Property Get ReportName() As String
    ReportName = Trim$(udtProps.ReportName)
End Property
Public Property Let UserName(Val As String)
'    If Not flgEditing Then Err.Raise 383
    
    udtProps.UserName = Val
    flgDirty = True
End Property
Public Property Get UserName() As String
    UserName = Trim$(udtProps.UserName)
End Property
Public Property Let Number(Val As String)
'    If Not flgEditing Then Err.Raise 383
    
'    objValid.RuleBroken "Number", (Len(Trim$(val)) = 0)
    udtProps.Number = Val
    flgDirty = True
End Property
Public Property Get Number() As String
    Number = Trim$(udtProps.Number)
End Property
Public Property Let DateOfError(Val As String)
 '   If Not flgEditing Then Err.Raise 383
    
    If (IsDate(Val)) Or (Val = "0") Then udtProps.DateOfError = Val
    flgDirty = True
End Property
Public Property Get DateOfError() As String
    If CLng(udtProps.DateOfError) = 0 Then
        DateOfError = ""
    Else
        DateOfError = Format((udtProps.DateOfError), "dd/mm/yyyy")
    End If
End Property
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property

Private Sub Save()
On Error GoTo ERR_Handler

Dim rs As ADODB.Recordset
'Dim cmd As ADODB.Command
Dim strSQL As String
Dim strTextLog As String
Dim iFilenum As Integer
Dim strMsg As String
Dim x As String
Dim fs As Scripting.FileSystemObject
    Set rs = New ADODB.Recordset
    strMsg = Trim$(udtProps.Description) & ", " & Trim$(udtProps.Number) & ", " & Trim$(udtProps.DateOfError) & ", " & Trim$(udtProps.FormName) & ", " & Trim$(udtProps.ReportName) & ", " & Trim$(udtProps.UserName)
    On Error Resume Next
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Trim$(udtProps.Number) = "Diagnostic" Then
        strTextLog = oPC.ServerRootPath & "\ErrorP.txt"
    Else
        strTextLog = oPC.ServerRootPath & "\ErrorLog.txt"
    End If
    iFilenum = FreeFile
    Open strTextLog For Append As #iFilenum
    Print #iFilenum, strMsg
    Close
    
EXIT_Handler:
    Set rs = Nothing
    Set K = Nothing
    Exit Sub
ERR_Handler:
    pResult = oPC.CO.Errors(0).NativeError
    GoTo EXIT_Handler
End Sub

Private Sub DeleteObject(ID As Long)
On Error GoTo ERR_Handler

  Dim strSQL As String
  Dim lngRecs As Long
  
    strSQL = "DELETE * FROM tblError WHERE Error_ID=" & ID
    CO.Execute strSQL
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    GoTo EXIT_Handler
End Sub

Private Sub Fetch(ID As Long)
  Dim rs As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim strSQL As String
  
  Set rs = New ADODB.Recordset
  Set cmd = New ADODB.Command
  strSQL = "SELECT * FROM tblError WHERE Error_ID=" & ID
  Set cmd.ActiveConnection = oPC.CO
  cmd.CommandText = strSQL
  rs.Open cmd, , adOpenDynamic, adLockReadOnly
  With rs
    udtProps.ID = .Fields("Error_ID")
    udtProps.Description = .Fields("ERR_Description")
    udtProps.DateOfError = .Fields("ERR_Date")
    udtProps.FormName = .Fields("ERR_Form")
    udtProps.ReportName = .Fields("ERR_Report")
    udtProps.UserName = .Fields("ERR_User")
    .Close
  End With
  Set rs = Nothing
  Set cmd = Nothing
End Sub
Public Sub SetError(pNumber As String, pDescription As String, pDateOfError As Date, pForm As String, pReport As String, pUser As String)
    If (pNumber = "Diagnostic" And bDiagnosticsOn = True) Or pNumber <> "Diagnostic" Then
        If Not IsEditing Then
            BeginEdit
        End If
        DateOfError = pDateOfError
        Description = pDescription
        Number = pNumber
        FormName = pForm
        ReportName = pReport
        UserName = pUser
        ApplyEdit
    End If
End Sub
