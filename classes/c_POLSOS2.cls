VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLSOS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objPersist As c_POLSOS2_P
Attribute objPersist.VB_VarHelpID = -1


Public Function LoadRecordsets(POLSOS As ADODB.Recordset, POLActions As ADODB.Recordset, Optional pETABefore As Date, Optional pTPID As Long, _
                                Optional pCOTPID As Long, Optional pPID As String, Optional pSTAFFID As Long, Optional pStatusChangeDate As Date)
    Set objPersist = New c_POLSOS2_P
        objPersist.Fetch POLSOS, POLActions, pETABefore, pTPID, pCOTPID, pPID, pSTAFFID, False, True, pStatusChangeDate
    Set objPersist = Nothing
End Function

Public Function LoadRecordsetsPIDSet(POLSOS As ADODB.Recordset, POLActions As ADODB.Recordset, Optional XMLArgs As String)
    Set objPersist = New c_POLSOS2_P
        objPersist.FetchPIDSet POLSOS, POLActions, XMLArgs
    Set objPersist = Nothing
End Function

