VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_SCL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, _
                        Optional plngTPID As Long, _
                        Optional pACNO As String, _
                        Optional pRef As String, _
                        Optional pPID As String, _
                        Optional pPCode As String, _
                        Optional pUNissued As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dSCLProps
Dim udtData As dSCLData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer

'    strDateCriteria = " TR_Date>= { d '" & ReverseDate(pdteDate1) & "'} AND TR_Date <= { d '" & ReverseDate(pdteDate2) & "'}"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
'    If plngTPID > 0 And Len(pPID) > 0 Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerPID2 WHERE P_ID = '" & pPID & "' AND R_TP_ID = " & plngTPID & " ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = True
'    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerPID2 WHERE P_Code = '" & pPCode & "' AND R_TP_ID = " & plngTPID & " ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = True
'    ElseIf pPCode > " " Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerPID2 WHERE  P_Code = '" & pPCode & "' OR P_EAN  = '" & pPCode & "' ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = False
'    ElseIf pRef > " " Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE  TR_Code = '" & pRef & "' OR R_APPROVALREF = '" & pRef & "' ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = False
'    ElseIf plngTPID <> 0 Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE " & strDateCriteria & " AND TR_TP_ID = " & plngTPID & " ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = False
'    ElseIf pACNO > " " Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE " & strDateCriteria & " AND  TP_ACNO = '" & pACNO & "' ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = True
'    ElseIf pdteDate1 >= #1/1/1995# Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE " & strDateCriteria & " ORDER BY TR_Date", enText, Array(), "", rs)
'        bIncludesLineDetails = False
'    ElseIf pUNissued Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM RPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
'        bIncludesLineDetails = False
'    Else
        If pUNissued Then
            lngResult = oBatch.RunGetRecordset("Select * FROM SCLPerTP_1 where DocStatus <> 4 ORDER BY DocDate", enText, Array(), "", rs)
        Else
            lngResult = oBatch.RunGetRecordset("Select * FROM SCLPerTP_1 ORDER BY DocDate", enText, Array(), "", rs)
        End If
        bIncludesLineDetails = False
'    End If
    
    Set objPB = New PropertyBag
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    lngCount = 0
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!ClaimID
            udtProps.TPID = FNN(rs!TPID)
            udtProps.DOCCode = FNS(rs!DOCCode)
            udtProps.DOCDate = FND(rs!DOCDate)
            udtProps.SupplierName = FNS(rs!SupplierName)
            udtProps.SupplierAcno = FNS(rs!SupplierAcno)
            udtProps.DocStatus = FNN(rs!DocStatus)
            udtProps.DocStatusF = FNS(rs!DocStatusF)
            udtProps.ClaimValue = FNDBL(rs!ClaimValue)
            udtProps.ClaimValueF = FNS(rs!ClaimValueF)
            udtProps.ClaimNeedsApproval = FNB(rs!ClaimNeedsApproval)
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




