VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_UTIL_P"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Status(pMsg As String)
Dim strMMsg As String
Public Enum DictionaryType
    dtInterestGroup = 1
    dtCustomerType = 2
    dtSupplierType = 3
    dtCategory = 4
    dtProductType = 5
End Enum
Function TableExists(pName As String) As Boolean
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
      
    cmd.CommandText = "TableExists"
    cmd.CommandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@pName", adVarChar, adParamInput, 50)
    cmd.Parameters.Append prm
    prm.Value = pName
    Set prm = cmd.CreateParameter("@pResult", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    cmd.execute
    TableExists = Trim(cmd.Parameters(1)) = 1
End Function
'Function DropTable(pName, pMsg As String, Optional pConn As ADODB.Connection) As Long
'    On Error GoTo ErrHandler
'Dim cmd As ADODB.Command
'Dim strSQL As String
'
'    If TableExists(pName) Then
'        Set cmd = New ADODB.Command
'        If pConn Is Nothing Then
'            cmd.ActiveConnection = oPC.COShort
'        Else
'            cmd.ActiveConnection = pConn
'        End If
'        cmd.CommandText = "DROP TABLE " & pName & ";"
'        cmd.CommandType = adCmdText
'        RaiseEvent Status(strMMsg & ": " & pMsg)
'        cmd.execute
'    End If
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_UTIL_P.DropTable(pName,pMsg,pConn)", Array(pName, pMsg, pConn)
'End Function
Function RunProc(pName As String, pParams As Variant, pMsg As String, Optional pConn As ADODB.Connection) As Long
On Error GoTo ERR_Handler
Dim cmd As New ADODB.Command
Dim retval
    
    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pName
    cmd.CommandType = adCmdStoredProc
    RaiseEvent Status(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            retval = cmd.execute(, pParams)
        Else
            cmd.execute
        End If
    Else
        cmd.execute
    End If
    
    Set cmd = Nothing
    
EXIT_Handler:
    Exit Function
ERR_Handler:
    If Err.Number = -2147217833 Then Resume Next
    Err.Raise vbError + 500
    GoTo EXIT_Handler
    Resume
End Function
Function RunSQL(pSQL As String) As Long
On Error GoTo ERR_Handler
Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = pSQL
    cmd.CommandType = adCmdText
    
    cmd.execute lngRecordsAffected
    Set cmd = Nothing
    RunSQL = lngRecordsAffected
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
ERR_Handler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Select Case Err
    Case -2147217900, -2147227904:
        Err.Raise vbObjectError + 555
        Resume EXIT_Handler
    Case Else
        Err.Raise vbError + 500
        Resume Next
        GoTo EXIT_Handler
End Select

End Function
Function RunGetRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset, Optional pConn As ADODB.Connection) As Long
On Error GoTo ERR_Handler


Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
    
    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pName
    cmd.CommandType = pType
    RaiseEvent Status(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            Set rs = cmd.execute(lngRecordsAffected, pParams)
        Else
            Set rs = cmd.execute(lngRecordsAffected)
        End If
    Else
        Set rs = cmd.execute(lngRecordsAffected)
    End If
    Set cmd = Nothing
    RunGetRecordset = lngRecordsAffected
EXIT_Handler:
    Exit Function
ERR_Handler:
    Err.Raise vbError + 500
    GoTo EXIT_Handler
    Resume
End Function

Function GetDynamicRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset, _
                    Optional pConn As ADODB.Connection)
    If pConn Is Nothing Then
        rs.open pName, oPC.COShort, adOpenDynamic, adLockOptimistic
    Else
        rs.open pName, pConn, adOpenDynamic, adLockOptimistic
    End If
End Function

Public Function AddDictRow(pVal As String, pSHort As String, pType As String)
On Error GoTo H
'    RunSQL "INSERT INTO tDICT (DICT_TYPE,DICT_DESC,DICT_SHORT,DICT_ACTIVE) VALUES (" & pType & ",'" & pVal & "','" & pSHort & "',1)"
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "INSERTDICTROW"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@TYPE", adVarChar, adParamInput, 10, pType)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@DESCRIPTION", adVarChar, adParamInput, 200, pVal)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@SHORTDESCRIPTION", adVarChar, adParamInput, 10, pSHort)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@ACTIVE", adBoolean, adParamInput, 10, True)
    cmd.Parameters.Append prm
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
H:
MsgBox "You cannot add this row. The or reported is: " & Error, vbExclamation, "Can't do this"
End Function
Public Function UpdateDictRow(pVal As String, pSHort As String, pType As String, PID As Long, Optional pActive As Boolean)
On Error GoTo H
    If Not IsNull(pActive) Then
        RunSQL "UPDATE tDICT SET DICT_DESC = '" & Replace(pVal, "'", "''") & "',DICT_Short = '" & pSHort & "',DICT_ACTIVE = " & IIf(pActive, 1, 0) & " WHERE DICT_ID = " & PID
    Else
        RunSQL "UPDATE tDICT SET DICT_DESC = '" & Replace(pVal, "'", "''") & "',DICT_Short = '" & pSHort & "' WHERE DICT_ID = " & PID
    End If
    Exit Function
H:
MsgBox Error
End Function
Public Function DeleteDictRow(PID As Long)
On Error GoTo H
    RunSQL "DELETE FROM tDICT WHERE DICT_ID = " & PID
    Exit Function
H:
MsgBox Error
End Function
Public Function RemoveStaffmember(pSMID As Long)
    'First look for staffmember called "Obsolete", if not found, create.
    'then set all SMIDs on transaction list from pSMID to that of "Obsolete"
    'then remove Staffmember pSMID from tStaffmember
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "REMOVESTAFFMEMBER"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@SMID", adInteger, adParamInput, , pSMID)
    cmd.Parameters.Append prm
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function

Public Function GetTextBite(pCode As String) As String
    On Error GoTo errHandler

Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "GetTextBite"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@Key", adVarChar, adParamInput, 25, Left(pCode, 25))
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@pResult", adVarChar, adParamOutput, 500)
    cmd.Parameters.Append prm
    
    cmd.execute
    GetTextBite = FNS(cmd.Parameters(1))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_UTIL_P.GetTextBite(pCode)", pCode
End Function
