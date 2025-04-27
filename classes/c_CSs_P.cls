VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CSs_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional pDocCOde As String, Optional pdteDate1 As Date, _
                Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As CSProps
Dim udtData As CSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
    pblnNoRecsReturned = False
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM CSPerPID WHERE P_ID = '" & pPID & "'  ORDER BY DocDate", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pDocCOde) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM CSPerTP WHERE TR_Code = '" & CLARG(pDocCOde) & "' ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM CSPerPID WHERE P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "' ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM CSPerTP WHERE  TRDate>= {d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= {d '" & ReverseDate(pdteDate2) & "'} ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM CSPerTP ORDER BY TR_Code", enText, Array(), "", rs)
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
    
    On Err GoTo errHandler
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!TransactionID
            udtProps.CS_GUID_ID = FNS(rs!CS_GUID)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DateIssued = FND(rs!TRDATE)
            udtProps.DateStarted = FND(rs!TR_CaptureDate)
            udtProps.Status = FNN(rs!TR_Status)
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs_P.Fetch(pblnNoRecsReturned,pDocCode,pdteDate1,pdteDate2,pPID,pPCode)", _
         Array(pblnNoRecsReturned, pDocCOde, pdteDate1, pdteDate2, pPID, pPCode)
End Function
Public Function FetchByZ_ID(ByRef pblnNoRecsReturned As Boolean, pZID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As CSProps
Dim udtData As CSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim OpenResult As Integer
Dim tmpStatus As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
    pblnNoRecsReturned = False
    lngResult = oBatch.RunGetRecordset("Select * FROM tCS   Left JOIN dbo.tCashup ON dbo.tCS.CS_GUID = dbo.tCashup.CU_XID  JOIN tTR ON CS_ID = TR_ID LEFT JOIN tStaffMember on CS_OperatorID = SM_ID WHERE CS_ZSessionID = '" & pZID & "' ORDER BY TR_Code", enText, Array(), "", rs)
    Set objPB = New PropertyBag
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        FetchByZ_ID = objPB.Contents
        Set objPB = Nothing
        GoTo EXIT_Handler
    End If
    
    On Err GoTo errHandler
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!TR_ID
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DateIssued = FND(rs!TR_DATE)
            udtProps.DateStarted = FND(rs!CS_StartSession)
            udtProps.DateIssued = FND(rs!CS_EndSession)
            udtProps.StaffName = FNS(rs!SM_Name)
            udtProps.CS_GUID_ID = FNS(rs!CS_GUID)
            udtProps.Status = FNN(rs!TR_Status)
            If FNN(rs!CU_ExplainedByStaffID) > 0 Then
                tmpStatus = 3
            ElseIf FNN(rs!CU_IssuedByStaffID) > 0 Then
                tmpStatus = 2
            Else
                tmpStatus = 0
            End If
            udtProps.Reportable = tmpStatus   'FNB(IIf(FNN(rs!CU_IssuedByStaffID) > 0, True, False))
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          FetchByZ_ID = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
    Set oBatch = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs_P.FetchByZ_ID(pblnNoRecsReturned,pZID)", Array(pblnNoRecsReturned, pZID)
End Function

