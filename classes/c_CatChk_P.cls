VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CatChk_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional CATCode As String, Optional CATShortCode As String, Optional pCode As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pSupplierID As Long, Optional pPEAN As String, Optional pPID As String) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dCatChkProps
Dim udtData As dCatChkData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim strDateLimits As String
Dim bIncludesLineDetails As Boolean
Dim OpenResult As Integer

    pblnNoRecsReturned = False
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strDateLimits = "TRDATE >= {d '" & ReverseDate(pdteDate1) & "'} AND TRDATE <= {d '" & ReverseDate(pdteDate2) & "'}"
'    If pSupplierID > 0 And Len(pPID) > 0 Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperTPID WHERE " & "P_ID = '" & pPID & "' AND TR_TP_ID = " & pSupplierID & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
'        bIncludesLineDetails = True
'    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM vCOSRperPID WHERE " & "P_Code = '" & pPID & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
'        bIncludesLineDetails = True
'    Else
    
    If Len(pCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperProduct WHERE P_Code = '" & CLARG(pCode) & "' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pPEAN) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperProduct WHERE P_EAN = '" & CLARG(pPEAN) & "' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
        
    ElseIf Len(CATCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperCat WHERE " & "CategoryName Like '" & CLARG(CATCode) & "%' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf Len(CATShortCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperCat WHERE " & "CategoryShort = '" & CLARG(CATShortCode) & "%' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
'    ElseIf plngTPID <> 0 Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM vCOSRperTP WHERE " & strDateLimits & " AND TR_TP_ID = " & plngTPID & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
'        bIncludesLineDetails = False
'    ElseIf pACNO > " " Then
'        lngResult = oBatch.RunGetRecordset("Select * FROM vCOSRperTP WHERE " & strDateLimits & " AND ACCNUM = '" & pACNO & "' ORDER BY TRDATE DESC", enText, Array(), "", rs)
'        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperCAT WHERE " & strDateLimits & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM vCATCHKperCAT  ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    
    End If
    Set objPB = New PropertyBag
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.CATCHKID = FNN(rs!CATCHKID)
            udtProps.CategoryName = FNS(rs!CategoryName)
            udtProps.CategoryCode = FNS(rs!CategoryShort)
            udtProps.DOCCode = FNS(rs!DOCCode)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.ProcessingDate = FND(rs!ProcessingDate)
            udtProps.OperatorName = FNS(rs!OperatorName)
            udtProps.OperatorShortName = FNS(rs!OperatorShortName)
            udtProps.SupervisorID = FNN(rs!CATCHK_Supervisor_ID)
            udtProps.SupervisorName = FNS(rs!SupervisorName)
            udtProps.SupervisorShortName = FNS(rs!SupervisorShortName)
            udtProps.Status = FNN(rs!DocStatus)
            If bIncludesLineDetails Then
            
            End If
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




