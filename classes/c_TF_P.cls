VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_TF_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional pDocCOde As String, Optional pdteDate1 As Date, _
                Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pUNissued As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As TFProps
Dim udtData As TFData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer
    
    strDateCriteria = "TRDate>= { d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= { d '" & ReverseDate(pdteDate2) & "'}"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
    pblnNoRecsReturned = False
    If Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM TFPerPID WHERE P_ID = '" & pPID & "'  ORDER BY DocDate", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pDocCOde) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM TFPerTP WHERE TR_Code = '" & CLARG(pDocCOde) & "' ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM TFPerPID WHERE P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "' ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM TFPerTP WHERE  TRDate>= {d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= {d '" & ReverseDate(pdteDate2) & "'} ORDER BY TRDate", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pUNissued Then
        lngResult = oBatch.RunGetRecordset("Select * FROM TFPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM TFPerTP ORDER BY TR_Code", enText, Array(), "", rs)
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
            udtProps.TRID = rs!TransactionID
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.Status = FNN(rs!TR_Status)
            udtProps.DestinationName = FNS(rs!STORE_Name)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.InOut = FNS(rs!TFR_INOUT)
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


