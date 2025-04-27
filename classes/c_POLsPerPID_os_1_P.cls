VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLsPerPID_os_1_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event lngProgress(p As Long)
Event lngMax(p As Long)
Dim iMaxRecords As Long
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

Public Function Fetch(pPID As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim udtProps As dPOLProps
Dim udtData As dPOLData
Dim pblnNoRecsReturned As Boolean
Dim rs2 As ADODB.Recordset
Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strAction As String
Dim bTmp As Boolean
Dim lngRC As Long
Dim iQtyRecs As Long
Dim i As Integer
Dim OpenResult As Integer
Dim strSQL As String
    strSQL = "SELECT * FROM vPOLsPerPID_os_1 WHERE P_ID = '" & pPID & "' Order BY P_Title "
    rs.CursorLocation = adUseClient

    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    iMaxRecords = rs.RecordCount
    
    Set objPB = New PropertyBag
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    i = 0

    Do While Not rs.eof
        With udtProps
            udtProps.Actions = IIf(CLng(rs!POL_LastActionDate) > 0, FNS(Format((rs!POL_LastActionDate), "dd-mm-yyyy")), "")
            udtProps.Ref = IIf(CLng(rs!POL_ETA) > 0, FNS(Format((rs!POL_ETA), "dd-mm-yyyy")), "")   'cheating here!
            udtProps.QtyReceivedSoFar = FNN(rs!POL_QTYReceivedSoFar)
            udtProps.PID = FNS(rs!P_ID)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TR_DATE)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyOutstanding = udtProps.QtySS + udtProps.QtyFirm - FNN(rs!POL_QTYReceivedSoFar)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.MainAuthor = FNS(rs!P_MainAuthor)
            udtProps.code = FNS(rs!P_Code)
            udtProps.EAN = FNS(rs!P_EAN)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            RaiseEvent lngProgress(lngCount)
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
        i = i + 1
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLsPerPID_os_1_P.Fetch(pPID)", pPID
End Function


Private Sub Class_Initialize()
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
End Sub

Private Sub Class_Terminate()
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

