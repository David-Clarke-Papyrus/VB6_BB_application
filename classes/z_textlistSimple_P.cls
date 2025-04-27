VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_textlistSimple_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum SimpleListType
    sltCNote = 1
    sltEntity = 2
    sltPublisher = 3
    sltAdhocQueries = 4
    sltFieldList = 5
    sltReportList = 6
End Enum

Public Function Fetch(ListName As SimpleListType, Optional pParam As String) As String
Dim rs As ADODB.Recordset
Dim objPB As PropertyBag
Dim udtProps As TextListSProps
Dim udtData As TextListSData
Dim lngCount As Long
Dim strSQL As String
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag
    On Error GoTo H
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Select Case ListName
        Case 1
                strSQL = "SELECT CNote FROM vUniqueCNotes WHERE ENTITY = '" & pParam & "' ORDER BY CNote"
        Case 2
            If pParam > "" Then
                strSQL = "SELECT Entity FROM vUniqueEntities WHERE Entity = '" & pParam & "' ORDER BY Entity"
            Else
                strSQL = "SELECT Entity FROM vUniqueEntities ORDER BY Entity"
            End If
        Case 3
            strSQL = "SELECT DISTINCT P_Publisher as PublisherName FROM PublisherList_TEMP ORDER BY P_Publisher"
        Case 4
            strSQL = "SELECT TABLE_NAME From PBKS.INFORMATION_SCHEMA.Views WHERE TABLE_NAME like 'ahv%' ORDER BY TABLE_NAME"
        Case 5
            strSQL = "SELECT COLUMN_NAME From INFORMATION_SCHEMA.COLUMNS  WHERE TABLE_NAME =  '" & pParam & "' ORDER BY COLUMN_NAME"
        Case 6
            strSQL = "SELECT RPT_NAME From tAdHocReports   ORDER BY RPT_NAME"
    End Select
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.eof
        If IsNull(rs(0)) Then GoTo Tail
        udtProps.Item = rs(0)
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
Tail:
        rs.MoveNext
    Loop
    rs.Close
    objPB.WriteProperty "Count", lngCount
    Fetch = objPB.Contents

    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    On Error Resume Next
  rs.Close
  Set rs = Nothing
  
  
  Err.Raise Err.Number
  Exit Function
  Resume
End Function

'#If conWin2000 Then
'
'Private Sub ObjectControl_Activate()
'  Set mobjContext = GetObjectContext
'End Sub
'
'Private Function ObjectControl_CanBePooled() As Boolean
'  ObjectControl_CanBePooled = False
'End Function
'
'Private Sub ObjectControl_Deactivate()
'  Set mobjContext = Nothing
'End Sub
'#EndIF

