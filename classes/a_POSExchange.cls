VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_POSExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub SendPOSExchange(pEXCHID As String, pOPSID As String, pZID As String)
    On Error GoTo errHandler
Dim Msg As String
Dim sFileName As String
Dim oShapeDB As New z_POSCLIConnectionShape
Dim sSQL As String
Dim tmprs As ADODB.Recordset
Dim oDE As New DataEnvironment1

    Check (oShapeDB.dbConnecttoShape = 0), ERR_GENERAL, "Failed to create database connection!"
    If Not rsZSession Is Nothing Then
        If rsZSession.State <> 0 Then
            rsZSession.Close
        End If
    End If
    sSQL = "SHAPE {SELECT 'E' as TYP,tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession APPEND (( SHAPE {SELECT EXCH_STATUS, EXCH_ID, EXCH_ZSESSION_ID,EXCH_OPSESSION_ID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE FROM tEXCHANGE WHERE EXCH_ID = '" & pEXCHID & "'}  AS POSExchange APPEND ({SELECT * FROM tCSL}  AS rsSALESLINES RELATE EXCH_ID TO CSL_EXCH_ID) AS SALESLINES,({SELECT * FROM tPayment}  AS rsPAYMENTS RELATE EXCH_ID TO PAY_EXCH_ID) AS PAYMENTS) AS POSExchange RELATE OPS_ID TO EXCH_OPSESSION_ID) AS POSExchange) AS OPSession RELATE Z_ID TO OPS_Z_ID) AS OPSession"
    Set rsZSession = Nothing
    Set rsZSession = New ADODB.Recordset
    rsZSession.Open sSQL, oShapeDB.DBConn, adOpenStatic
    Set rsZSession.ActiveConnection = Nothing
    
    If Not rsZSession.EOF Then
        sFileName = oPC.TillCode & "-" & Format(Now(), "DDHHNNSS") 'Format(oGD.GetNextFileNum(), "00000")
        sFileName = "\" & sFileName & ".pos"
        rsZSession.Save sOutBox & sFileName, adPersistADTG
    End If
    If pEXCHID > "" Then
        SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
        oPC.DBConn.Execute SQL
    End If
    oShapeDB.dbCloseConnectShape
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSExchange.SendPOSExchange"
End Sub

Private Sub Class_Initialize()

End Sub
