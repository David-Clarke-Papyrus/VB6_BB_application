VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type orProps
    ID As Long
    Description As String * 250
    Number As String * 100
    DateOfor As Date
    UserName As String * 100
    FormName As String * 200
    ReportName As String * 200
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Dim udtProps As orProps
Dim udtSave As orProps
Dim bDiagnosticsOn As Boolean
Event Valid(IsValid As Boolean)
Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

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
    Exit Sub
    Resume
End Sub
Private Sub Class_Initialize()
    Set ooR = Me
    flgNew = True
End Sub

'Public Property Get IsValid() As Boolean
'    IsValid = (mobjValid.Count = 0)
'End Property
Public Sub SetDiagnostics(val As Boolean)
    bDiagnosticsOn = val
End Sub
Public Property Get DiagnosticsOn() As Boolean
    DiagnosticsOn = bDiagnosticsOn
End Property

Private Sub Class_Terminate()
    Set mobjValid = Nothing
End Sub

'Private Sub mobjValid_BrokenRule(ors As String)
'    RaiseEvent Valid(False)
'End Sub
'
'Private Sub mobjValid_NoBrokenRules()
'    RaiseEvent Valid(True)
'End Sub

Public Sub Load(ID As Long)
'    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    
    Fetch ID
'    mobjValid.RuleBroken "Description", False
'    mobjValid.RuleBroken "Number", False
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
Public Property Let Description(val As String)
'    If Not flgEditing Then Err.Raise 383
    
'    mobjValid.RuleBroken "Description", (Len(Trim$(val)) = 0)
    udtProps.Description = val
    flgDirty = True
End Property
Public Property Get Description() As String
    Description = Trim$(udtProps.Description)
End Property
Public Property Let FormName(val As String)
'    If Not flgEditing Then Err.Raise 383
    
    udtProps.FormName = val
    flgDirty = True
End Property
Public Property Get FormName() As String
    FormName = Trim$(udtProps.FormName)
End Property
Public Property Let ReportName(val As String)
'    If Not flgEditing Then Err.Raise 383
    
    udtProps.ReportName = val
    flgDirty = True
End Property
Public Property Get ReportName() As String
    ReportName = Trim$(udtProps.ReportName)
End Property
Public Property Let UserName(val As String)
'    If Not flgEditing Then Err.Raise 383
    
    udtProps.UserName = val
    flgDirty = True
End Property
Public Property Get UserName() As String
    UserName = Trim$(udtProps.UserName)
End Property
Public Property Let Number(val As String)
'    If Not flgEditing Then Err.Raise 383
    
'    mobjValid.RuleBroken "Number", (Len(Trim$(val)) = 0)
    udtProps.Number = val
    flgDirty = True
End Property
Public Property Get Number() As String
    Number = Trim$(udtProps.Number)
End Property
Public Property Let DateOfor(val As String)
 '   If Not flgEditing Then Err.Raise 383
    
    If (IsDate(val)) Or (val = "0") Then udtProps.DateOfor = val
    flgDirty = True
End Property
Public Property Get DateOfor() As String
    If CLng(udtProps.DateOfor) = 0 Then
        DateOfor = ""
    Else
        DateOfor = Trim$(udtProps.DateOfor)
    End If
End Property
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property
Private Property Get NameOforPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOforPC = Left(MachineName, NameSize)
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
    strMsg = Trim$(NameOforPC) & ": " & Trim$(udtProps.Description) & ", " & Trim$(udtProps.Number) & ", " & Trim$(udtProps.DateOfor) & ", " & Trim$(udtProps.FormName) & ", " & Trim$(udtProps.ReportName) & ", " & Trim$(udtProps.UserName)
    On Error Resume Next
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Trim$(udtProps.Number) = "Diagnostic" Then
        strTextLog = App.Path & "\orP.txt"
    Else
        strTextLog = oPC.SharedFolderRoot & "\orLog.txt"
    End If
    iFilenum = FreeFile
    Open strTextLog For Append As #iFilenum
    Print #iFilenum, strMsg
    Close
    
EXIT_Handler:
    Set rs = Nothing
    Set k = Nothing
    Exit Sub
ERR_Handler:
    pResult = oPC.COShort.oRS(0).Nativeor
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub DeleteObject(ID As Long)
On Error GoTo ERR_Handler

  Dim strSQL As String
  Dim lngRecs As Long
  
    strSQL = "DELETE * FROM tblor WHERE or_ID=" & ID
    CO.execute strSQL
    
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
  strSQL = "SELECT * FROM tblor WHERE or_ID=" & ID
  Set cmd.ActiveConnection = oPC.COShort
  cmd.CommandText = strSQL
  rs.open cmd, , adOpenDynamic, adLockReadOnly
  With rs
    udtProps.ID = .fields("or_ID")
    udtProps.Description = .fields("_Description")
    udtProps.DateOfor = .fields("_Date")
    udtProps.FormName = .fields("_Form")
    udtProps.ReportName = .fields("_Report")
    udtProps.UserName = .fields("_User")
    .Close
  End With
  Set rs = Nothing
  Set cmd = Nothing
End Sub
Public Sub SetError(pNumber As String, pDescription As String, pDateOfor As Date, pForm As String, pReport As String, pUser As String)
    If (pNumber = "Diagnostic" And bDiagnosticsOn = True) Or pNumber <> "Diagnostic" Then
        If Not IsEditing Then
            BeginEdit
        End If
        DateOfor = pDateOfor
        Description = pDescription
        Number = pNumber
        FormName = pForm
        ReportName = pReport
        UserName = pUser
        ApplyEdit
    End If
End Sub
