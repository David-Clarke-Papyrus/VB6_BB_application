Attribute VB_Name = "mExportImport"
Option Explicit

Enum enumIETypes
   EXPORTCUSTOMERS
   IMPORTCUSTOMERS
   EXPORTDEBTORSTRADING
   EXPORTCREDITORSTRADING
End Enum
Function ConvertIETypes(pIn As enumIETypes) As String
    Select Case pIn
        Case EXPORTCUSTOMERS
            ConvertIETypes = "ExportCustomers"
        Case IMPORTCUSTOMERS
            ConvertIETypes = "ImportCustomers"
        Case EXPORTDEBTORSTRADING
            ConvertIETypes = "ExportDebtorsTrading"
        Case EXPORTCREDITORSTRADING
            ConvertIETypes = "ExportCreditorsTrading"
            
        Case Else
            ConvertIETypes = "Unknown"
    End Select

End Function


Public Function GetLastDate(pType As enumIETypes) As Date
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "GetLastIEDATE"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@IETYPE", adVarChar, adParamInput, 50, ConvertIETypes(pType))
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@IEDATE", adDate, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.Execute
    GetLastDate = CDate(cmd.Parameters(1).Value)

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
Public Function GetLastTRID(pType As enumIETypes) As Long
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "GetLastIETR"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@IETYPE", adVarChar, adParamInput, 50, ConvertIETypes(pType))
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@IETRID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.Execute
    GetLastTRID = CLng(cmd.Parameters(1).Value)

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

