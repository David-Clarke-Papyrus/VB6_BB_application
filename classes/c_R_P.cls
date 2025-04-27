VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_R_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, _
                        Optional plngTPID As Long, _
                        Optional pACNO As String, _
                        Optional pRef As String, _
                        Optional pdteDate1 As Date, _
                        Optional pdteDate2 As Date, _
                        Optional pPID As String, _
                        Optional pPCode As String, _
                        Optional pUNissued As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dRProps
Dim udtData As dRData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer

    strDateCriteria = " TR_Date>= { d '" & ReverseDate(pdteDate1) & "'} AND TR_Date <= { d '" & ReverseDate(pdteDate2) & "'}"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If plngTPID > 0 And Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerPID2 WHERE P_ID = '" & pPID & "' AND R_TP_ID = " & plngTPID & " ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerPID2 WHERE P_Code = '" & CLARG(pPCode) & "' AND R_TP_ID = " & plngTPID & " ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf pPCode > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerPID2 WHERE  P_Code = '" & CLARG(pPCode) & "' OR P_EAN  = '" & CLARG(pPCode) & "' ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pRef > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE  TR_Code = '" & CLARG(pRef) & "' OR R_APPROVALREF = '" & CLARG(pRef) & "' ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf plngTPID <> 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE " & strDateCriteria & " AND TR_TP_ID = " & plngTPID & " ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pACNO > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE " & strDateCriteria & " AND  TP_ACNO = '" & CLARG(pACNO) & "' ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE " & strDateCriteria & " ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pUNissued Then
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP ORDER BY TR_Date", enText, Array(), "", rs)
        bIncludesLineDetails = False
    End If
    
    Set objPB = New PropertyBag
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!TR_ID
            udtProps.TPNAME = FNS(rs!TP_Name)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.TPID = FNN(rs!TR_TP_ID)
            udtProps.DOCDate = FND(rs!TR_DATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.Status = FNN(rs!TR_Status)
            udtProps.ReturnStatus = FNN(rs!R_Status)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.RType = FNS(rs!R_Type)
            udtProps.ApprovalRef = FNS(rs!R_ApprovalRef)
            udtProps.ApprovalTermDate = FND(rs!R_ApprovalTermDate)
'            If bIncludesLineDetails Then
'                udtProps.QtySystem = FNN(rs!RL_QtySystem)
'                udtProps.QtyRequested = FNN(rs!RL_QtyRequested)
'                udtProps.QtyApproved = FNN(rs!RL_QtyApproved)
'                udtProps.QtyReturned = FNDBL(rs!RL_QtyReturned)
'                udtProps.QtyCounted = FNDBL(rs!RL_QtyCounted)
'            End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set oBatch = Nothing
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function


