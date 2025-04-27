VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_App_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional pTRCode As String, Optional plngTPID As Long, Optional pACNO As String, Optional pRef As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pUNissued As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim objDisplay As d_APP
Dim dteDate1 As Date
Dim udtProps As APPProps
Dim udtData As APPData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer

    strDateCriteria = " TRDate>= {d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= {d '" & ReverseDate(pdteDate2) & "'}"
    On Error GoTo ERR_Handler
    
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If plngTPID > 0 And Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerPID WHERE P_ID = '" & pPID & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerPID WHERE P_Code = '" & CLARG(pPID) & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf pPCode > "" Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerPID WHERE  P_Code = '" & CLARG(pPCode) & "' or P_EAN = '" & pPCode & "' ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pTRCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerTP WHERE TR_Code = '" & CLARG(pTRCode) & "' ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf plngTPID <> 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerTP WHERE " & strDateCriteria & " AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pACNO > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerTP WHERE " & strDateCriteria & " AND ACCNUM = '" & CLARG(pACNO) & "' ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerTP WHERE " & strDateCriteria & " ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pUNissued Then
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM APPsPerTP ORDER BY TRDate", enText, Array(), "", rs)
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
    
    
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!TransactionID
            udtProps.TPNAME = FNS(rs!TPNAME)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.Status = rs!TR_Status
            udtProps.CustomerDisplay = FNS(rs!CustomerDisplay)
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
    Set objDisplay = Nothing
    Exit Function
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function


