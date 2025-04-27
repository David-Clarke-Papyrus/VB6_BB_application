VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Payments_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional plngTPID As Long, Optional pACNO As String, Optional pDocCOde As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pRef As String, Optional rsR As ADODB.Recordset) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dPaymentProps
Dim udtData As dCOData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim strDateCriteria As String
Dim OpenResult As Integer
    
    strDateCriteria = "DepositDate>= {d '" & ReverseDate(pdteDate1) & "'} AND DepositDate <= {d '" & ReverseDate(pdteDate2) & "'}"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set oBatch = New z_SQL
    pblnNoRecsReturned = False
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If Len(pDocCOde) > 0 Then
        rs.open "Select * FROM PaymentsPerCustomer WHERE DepositCode = '" & pDocCOde & "' OR PaymentCode = '" & pDocCOde & "' ORDER BY DepositDate", oPC.COShort, adOpenStatic, adLockOptimistic
        bIncludesLineDetails = False
    ElseIf plngTPID <> 0 Then
        rs.open "Select * FROM PaymentsPerCustomer WHERE " & strDateCriteria & " AND (DepositorID = " & plngTPID & " OR CustomerID = " & plngTPID & ")  ORDER BY DepositDate", oPC.COShort, adOpenStatic, adLockOptimistic
        bIncludesLineDetails = False
    ElseIf pACNO > " " Then
        rs.open "Select * FROM PaymentsPerCustomer WHERE " & strDateCriteria & " AND DepositorAcno = '" & pACNO & "' OR CustomerAcno = '" & pACNO & "' ORDER BY DepositDate", oPC.COShort, adOpenStatic, adLockOptimistic
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        rs.open "Select * FROM PaymentsPerCustomer WHERE " & strDateCriteria & " ORDER BY DepositDate DESC", oPC.COShort, adOpenStatic, adLockOptimistic
        bIncludesLineDetails = False
    Else
        rs.open "Select * FROM PaymentsPerCustomer  ORDER BY DepositDate", oPC.COShort, adOpenStatic, adLockOptimistic
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
            udtProps.DepositID = rs!DepositID
            udtProps.DepositorName = FNS(rs!DepositorName)
            udtProps.DepositorFullName = FNS(rs!DepositorFullName)
            udtProps.CustomerName = FNS(rs!CustomerName)
            udtProps.CustomerFullName = FNS(rs!CustomerFullName)
            udtProps.DepositCode = FNS(rs!DepositCode)
            udtProps.DepositDate = FND(rs!DepositDate)
            udtProps.CaptureDate = FND(rs!DepositDate)
            udtProps.DepositorID = FNN(rs!DepositorID)
            udtProps.DepositorAcNo = FNS(rs!DepositorAcNo)
            udtProps.CustomerAcno = FNS(rs!CustomerAcno)
            udtProps.RemittanceReference = FNS(rs!RemittanceReference)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.Status = FNN(rs!TR_Status)
            udtProps.DepositAmount = FNDBL(rs!DepositAmount)
            udtProps.DepositSettlementDiscount = FNDBL(rs!DepositSettlementAmount)
         
         '   udtProps.DepositType = FNS(rs!Type)
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
    Set rsR = rs
    rs.ActiveConnection = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set oBatch = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Payments_P.Fetch(pblnNoRecsReturned,plngTPID,pACNO,pDocCode,pdteDate1,pdteDate2,pRef," & _
        "rsR)", Array(pblnNoRecsReturned, plngTPID, pACNO, pDocCOde, pdteDate1, pdteDate2, pRef, rsR)
End Function



