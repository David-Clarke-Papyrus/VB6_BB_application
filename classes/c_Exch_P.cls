VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Exch_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, _
                        Optional plngTPID As Long, _
                        Optional pACNO As String, _
                        Optional pExchangeNum As Long, _
                        Optional pRef As String, _
                        Optional pdteDate1 As Date, _
                        Optional pdteDate2 As Date, _
                        Optional pPID As String, _
                        Optional pPCode As String, _
                        Optional pPIID As Long) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dPOProps
Dim udtData As dPOData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer

    strDateCriteria = "TRDate>= { d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= { d '" & ReverseDate(pdteDate2) & "'}"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If plngTPID > 0 And Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerPID WHERE P_ID = '" & pPID & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And pPIID > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerPID WHERE POL_PI_ID = '" & pPIID & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerPID WHERE P_Code = '" & CLARG(pPCode) & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf pPCode > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerPID WHERE  P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "' ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pRef) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerTP WHERE TR_Code = '" & CLARG(pRef) & "' ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf plngTPID <> 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerTP WHERE " & strDateCriteria & " AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pACNO > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerTP WHERE " & strDateCriteria & " AND TR_Code = '" & CLARG(pACNO) & "' ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerTP WHERE " & strDateCriteria & " ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM EXCHsPerTP ORDER BY TRDate DESC", enText, Array(), "", rs)
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
            udtProps.TPNAME = FNS(rs!TPNAME)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.Status = FNN(rs!TR_Status)
            udtProps.TPAccNo = FNS(rs!Accnum)
            udtProps.Status = FNS(rs!TR_Status)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.Log = Left(FNS(rs!TR_Log), 300)
'            If bIncludesLineDetails Then
'              '  udtProps.PID = FNN(rs!POL_P_ID)
'                udtProps.Price = FNN(rs!POL_Price)
'                udtProps.QtySS = FNN(rs!POL_QtySS)
'                udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
'                udtProps.Discount = FNDBL(rs!POL_Discount)
'                udtProps.VATRate = FNDBL(rs!POL_VATRATE)
'          '      udtProps.POLID = FNN(rs!POL_ID) 'See notes on correspoinding c_CO_P
'            End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
            If lngCount > MAXBROWSE Then Exit Do
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


