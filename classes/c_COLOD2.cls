VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLOD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objPersist As c_COLsOD_P2
Attribute objPersist.VB_VarHelpID = -1


Public Function LoadRecordsets(COLS As ADODB.Recordset, POLS As ADODB.Recordset, COLActs As ADODB.Recordset, Optional pETABefore As Date, _
    Optional pTPID As Long, Optional pPID As String, Optional pSTAFFID As Long, Optional pStatusChangeDate As Date)
    Set objPersist = New c_COLsOD_P2
        objPersist.Fetch COLS, POLS, COLActs, pETABefore, pTPID, pPID, pSTAFFID, pStatusChangeDate
    Set objPersist = Nothing
End Function


