VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COs_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, pOrderType As enumCOType, Optional plngTPID As Long, Optional pACNO As String, Optional pDocCOde As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String, _
                Optional pPIID As Long, Optional pRef As String, Optional pOrderLineRef As String, Optional pUNissued As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dCOProps
Dim udtData As dCOData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer
    
    strDateCriteria = "TRDate>= {d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= {d '" & ReverseDate(pdteDate2) & "'}"
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
    pblnNoRecsReturned = False
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If plngTPID <> 0 And Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerPID WHERE CO_OrderType = " & pOrderType & " AND P_ID = '" & pPID & "' AND  TR_TP_ID = " & plngTPID & " ORDER BY DocDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And pPIID > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerPID WHERE   CO_OrderType = " & pOrderType & " AND COL_PI_ID = " & pPIID & " AND TR_TP_ID = " & plngTPID & " ORDER BY DocDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerPID WHERE   CO_OrderType = " & pOrderType & " AND P_Code = '" & CLARG(pPCode) & "' AND TR_TP_ID = " & plngTPID & " ORDER BY DocDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerPID WHERE   CO_OrderType = " & pOrderType & " AND P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "'  ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pOrderLineRef) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select DISTINCT TOP (100) PERCENT TP_ID, TRDate, TransactionID, TPName, TPPhone, AccNum, TR_Code, TR_NOTE, TR_Status, ADD_BusPhone, ADD_Fax, " _
                   & " TR_TP_ID , TR_CaptureDate, CO_OrderNUm, TP_Initials, TP_Title, CO_OrderType, TR_STAFFID, CustomerDisplay " _
                   & " FROM COsPerPID WHERE   CO_OrderType = " & pOrderType & " AND (OrderLineRef = '" & CLARG(pOrderLineRef) & "' OR TPName Like '%" & CLARG(pOrderLineRef) & "%')  ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pDocCOde) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE  CO_OrderType = " & pOrderType & " AND TR_Code = '" & CLARG(pDocCOde) & "' ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf plngTPID <> 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE  CO_OrderType = " & pOrderType & " AND TR_TP_ID = " & plngTPID & " ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pACNO > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE  CO_OrderType = " & pOrderType & " AND ACCNUM = '" & CLARG(pACNO) & "' ORDER BY TR_Code", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pRef > "" Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE  CO_OrderType = " & pOrderType & " AND  CO_OrderNUM = '" & CLARG(pRef) & "' ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE  CO_OrderType = " & pOrderType & " AND " & strDateCriteria & " ORDER BY TRDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pUNissued Then
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    Else
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP WHERE  CO_OrderType = " & pOrderType & " ORDER BY TR_Code", enText, Array(), "", rs)
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
            udtProps.TPNAME = FNS(rs!TPNAME) & IIf(rs!TP_Title > "", ", " & rs!TP_Title, "") & IIf(rs!TP_Initials > "", " " & rs!TP_Initials, "")
            udtProps.CustomerDisplay = FNS(rs!CustomerDisplay)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.TPAccNo = FNS(rs!Accnum)
            udtProps.OrderType = FNN(rs!CO_OrderType)
            udtProps.StaffID = FNN(rs!TR_StaffID)
'            If bIncludesLineDetails Then
'                udtProps.Ref = FNS(rs!COL_Ref)
'                udtProps.Deposit = FNN(rs!COL_Deposit)
'                udtProps.Discount = FNDBL(rs!COL_DiscountPercent)
'                udtProps.Price = FNN(rs!COL_Price)
'               ' udtProps.COLID = FNN(rs!COL_ID)  ' 17/2/2004 : removed this field from query as it causes duplicates in frmBrowseCOs, also set query to DISTINCT
'                udtProps.PID = FNS(rs!COL_P_ID)
'                If Len(rs!COL_P_ID) > 0 Or pPCode > "" Or pPIID > 0 Then
'                    udtProps.Qty = FNN(rs!COL_Qty)
'                End If
'            End If
            udtProps.Status = FNN(rs!TR_Status)
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


Public Function FetchOrdersToInvoice(ByRef pblnNoRecsReturned As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dCOProps
Dim udtData As dCOData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer
    
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
    pblnNoRecsReturned = False
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        
        lngResult = oBatch.RunGetRecordset("Select * FROM COsPerTP  JOIN vGDNsToInvoice ON TransactionID = TR_ID WHERE  CO_OrderType = 1   ORDER BY TR_Code", enText, Array(), "", rs)
    
    Set objPB = New PropertyBag
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        FetchOrdersToInvoice = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!TransactionID
            udtProps.TPNAME = FNS(rs!TPNAME) & IIf(rs!TP_Title > "", ", " & rs!TP_Title, "") & IIf(rs!TP_Initials > "", " " & rs!TP_Initials, "")
            udtProps.CustomerDisplay = FNS(rs!CustomerDisplay)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.TPAccNo = FNS(rs!Accnum)
            udtProps.OrderType = FNN(rs!CO_OrderType)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.Status = FNN(rs!TR_Status)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
            If lngCount > MAXBROWSE Then Exit Do
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          FetchOrdersToInvoice = .Contents
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


