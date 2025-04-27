VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_APPLsPerTPPID_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pTPID As Long, pPID As String, pPCode As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim udtProps As dAPPLProps
Dim udtData As dAPPLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    If pTPID > 0 Then
        If pPID > "" Then
            strSQL = "SELECT * FROM APPLsPerTPPID WHERE TP_ID = " & pTPID & " AND P_ID = '" & pPID & "'"
        Else
            If pPCode > "" Then
                strSQL = "SELECT * FROM APPLsPerTPPID WHERE TP_ID = " & pTPID & " AND P_Code = '" & pPCode & "'"
            Else
                strSQL = "SELECT * FROM APPLsPerTPPID WHERE TP_ID = " & pTPID
            End If
        End If
    Else
        strSQL = "SELECT * FROM APPLsPerTPPID WHERE P_ID = '" & pPID & "'"
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    Do While Not rs.eof
        With udtProps
            udtProps.APPLID = FNS(rs!APPL_ID)
            udtProps.Qty = FNN(rs!APPL_Qty) - FNN(rs!APPL_QtyReturned)
            udtProps.DiscountRate = FNDBL(rs!APPL_DiscountRate)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.TRID = FNN(rs!TransactionID)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.code = FNS(rs!CodeF)
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_APPLsPerTPPID_P.Fetch(pTPID,pPID,pPCode)", Array(pTPID, pPID, pPCode)
End Function




