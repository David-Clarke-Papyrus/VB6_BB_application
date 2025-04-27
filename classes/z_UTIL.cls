VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_UTIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents oUtilPersist As z_UTIL_P
Attribute oUtilPersist.VB_VarHelpID = -1
Event Status(pMsg As String)
Function TableExists(pName As String) As Boolean
Dim bResult As Boolean
    Set oUtilPersist = New z_UTIL_P
    TableExists = oUtilPersist.TableExists(pName)
    Set oUtilPersist = Nothing
End Function

Sub oUtilPersist_Status(pMsg As String)
    RaiseEvent Status(pMsg)
End Sub
'Function DropTable(pName, pMsg As String, Optional pConn As ADODB.Connection) As Long
'
'    Set oUtilPersist = New z_UTIL_P
'    oUtilPersist.DropTable pName, pMsg
'    Set oUtilPersist = Nothing
'
'End Function
Function RunProc(pName As String, pParams As Variant, pMsg As String, Optional pConn As ADODB.Connection) As Long
    
    Set oUtilPersist = New z_UTIL_P
    oUtilPersist.RunProc pName, pParams, pMsg
    Set oUtilPersist = Nothing
End Function
Function RunSQL(pSQL As String, Optional pConn As ADODB.Connection) As Long
    
    Set oUtilPersist = New z_UTIL_P
    oUtilPersist.RunSQL pSQL
    Set oUtilPersist = Nothing


End Function
Public Function AddDictRow(pVal As String, pSHort As String, pType As String)
    Set oUtilPersist = New z_UTIL_P
    oUtilPersist.AddDictRow pVal, pSHort, pType
    Set oUtilPersist = Nothing
End Function
Public Function UpdateDictRow(pVal As String, pSHort As String, pType As String, PID As Long, Optional pActive As Boolean)
    Set oUtilPersist = New z_UTIL_P
    oUtilPersist.UpdateDictRow pVal, pSHort, pType, PID, pActive
    Set oUtilPersist = Nothing
End Function
Public Function DeleteDictRow(PID As Long)
    Set oUtilPersist = New z_UTIL_P
    oUtilPersist.DeleteDictRow PID
    Set oUtilPersist = Nothing
End Function
Public Function RemoveStaffmember(pSMID As Long)
    Set oUtilPersist = New z_UTIL_P
    oUtilPersist.RemoveStaffmember pSMID
    Set oUtilPersist = Nothing
End Function
Public Function GetTextBite(pCode As String) As String
    Set oUtilPersist = New z_UTIL_P
    GetTextBite = oUtilPersist.GetTextBite(pCode)
    Set oUtilPersist = Nothing
End Function
'GetTextBite
