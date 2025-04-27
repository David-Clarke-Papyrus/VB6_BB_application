VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ro_POL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pTRID As Long) As String
Dim lngCount As Long
Dim udtProps As dPOLProps
Dim udtData As dPOLData
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
'Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim strAction As String
Dim bTmp As Boolean
Dim lngRC As Long

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
   ' strSQL = "SELECT *,ISNULL(POL_FULFILLED,'OS') as POLFulfilled,dbo.CodeF(P_CODE,P_EAN,0) as CodeF FROM tPOL JOIN tPRODUCT ON POL_P_ID = P_ID WHERE POL_TR_ID = " & pTRID
    strSQL = "SELECT * FROM POL_Fetch WHERE POL_TR_ID=" & pTRID
    rs.CursorLocation = adUseClient
    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    If rs.eof Then
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Set rs2 = New ADODB.Recordset
    Do While Not rs.eof
        With udtProps
            rs2.open "SELECT * FROM tPOLA WHERE POLA_POL_ID = " & CLng(rs!POL_ID) & " ORDER BY POLA_ID DESC", oPC.COShort, adOpenForwardOnly
            strAction = ""
'            If FNS(rs!POL_LastAction) > "" Then
'                strAction = Format(rs!POL_LastActionDate, "dd/mm/yyyy") & ": " & FNS(rs!POL_LastAction)
'            Else
'                strAction = ""
'            End If
            Do While Not rs2.eof
                strAction = strAction & IIf(Len(strAction) > 0, vbCrLf, "") & Format(rs2!POLA_Date, "dd/mm/yyyy") & ": " & ConvertPOLActionCodes(rs2!POLA_Code, bTmp)
                rs2.MoveNext
            Loop
            rs2.Close
            udtProps.Actions = strAction
            udtProps.Ref = FNS(rs!POL_REF)
            udtProps.Price = FNN(rs!POL_Price)
            udtProps.ForeignPrice = FNN(rs!POL_ForeignPrice)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyReceivedSoFar = FNN(rs!POL_QTYReceivedSoFar)
            udtProps.Discount = FNDBL(rs!POL_Discount)
            udtProps.POLID = FNN(rs!POL_ID)
            udtProps.PID = FNS(rs!POL_P_ID)
           ' udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.Fulfilled = IIf(FNS(rs.fields("POL_FulFilled")) = "", "OS", FNS(rs.fields("POL_FulFilled")))
           ' udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyOutstanding = udtProps.QtySS + udtProps.QtyFirm - FNN(rs!POL_QTYReceivedSoFar)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.MainAuthor = FNS(rs!P_MainAuthor)
            udtProps.code = FNS(rs!P_Code)
            udtProps.CodeF = FNS(rs!CodeF)
            udtProps.EAN = FNS(rs!P_EAN)
            udtProps.ETA = FND(rs!POL_ETA)
            udtProps.Replacementfor = FND(rs!POL_REPLACEMENTFOR)
            udtProps.VATRate = FNDBL(rs.fields("POL_VATRATE"))
            udtProps.LastAction = FNS(rs.fields("POL_LastAction"))
            udtProps.LastActionDate = FND(rs.fields("POL_LastActionDate"))
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
    Set rs2 = Nothing
    rs.Close
    Set rs = Nothing
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function




