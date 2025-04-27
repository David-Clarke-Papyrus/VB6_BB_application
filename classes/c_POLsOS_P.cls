VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLsOS_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event lngProgress(p As Long)
Event lngMax(p As Long)
Dim iMaxRecords As Long
Dim rs As ADODB.Recordset

Public Function PrepareToFetch(TotalRecs As Long, Optional pETABefore As Date, Optional TPID As Long, Optional pPID As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer

    If pETABefore <> CDate(0) Then
        If TPID > 0 Then
            If pPID > "" Then
                strSQL = "SELECT * FROM POLsOS WHERE TPID = " & TPID & " AND POL_ETA < {d '" & ReverseDate(pETABefore) & "' }" & " AND POL_P_ID = '" & pPID & "'  ORDER BY TP_Name"
            Else
                strSQL = "SELECT * FROM POLsOS WHERE TPID = " & TPID & " AND POL_ETA < {d '" & ReverseDate(pETABefore) & "' }  ORDER BY TP_Name"
            End If
        Else
            If pPID > "" Then
                strSQL = "SELECT * FROM POLsOS WHERE POL_ETA < {d '" & ReverseDate(pETABefore) & "'}" & " AND POL_P_ID = '" & pPID & "'  ORDER BY TP_Name"
            Else
                strSQL = "SELECT * FROM POLsOS WHERE POL_ETA < {d '" & ReverseDate(pETABefore) & "'} ORDER BY TP_Name"
            End If
        End If
    Else
        If TPID > 0 Then
            If pPID > "" Then
                strSQL = "SELECT * FROM POLsOS WHERE TPID = " & TPID & " AND POL_P_ID = '" & pPID & "'  ORDER BY TP_Name"
            Else
                strSQL = "SELECT * FROM POLsOS WHERE TPID = " & TPID & "  ORDER BY TP_Name"
            End If
        Else
            If pPID > "" Then
                strSQL = "SELECT * FROM POLsOS" & " WHERE POL_P_ID = '" & pPID & "'  ORDER BY TP_Name"
            Else
                strSQL = "SELECT * FROM POLsOS" & "  ORDER BY TP_Name"
            End If
        End If
    End If
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    iMaxRecords = rs.RecordCount
    TotalRecs = iMaxRecords
    RaiseEvent lngMax(iMaxRecords)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLsOS_P.PrepareToFetch(TotalRecs,pETABefore,TPID,pPID)", Array(TotalRecs, pETABefore, _
         TPID, pPID)
End Function

Public Function Fetch(Optional pETABefore As Date, Optional pTPID As Long, Optional pCOTPID As Long, Optional pPID As String, Optional pSTAFFID As Long, _
                        Optional pSubstitutions As Boolean, Optional pSkipactions As Boolean, Optional POLSetXML As String) As String
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
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer

    Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    If POLSetXML > "" Then
        cmd.CommandText = "vGetPOLSOS_perPO"
        Set prm = cmd.CreateParameter("@XMLArgs", adVarChar, adParamInput, 3000, POLSetXML)
        cmd.Parameters.Append prm
    Else
        If pSubstitutions = True Then
            If pPID > "" Then
                cmd.CommandText = "vGetPOLSOSSubs"
            Else
                cmd.CommandText = "vGetPOLSOSSubs2"
            End If
        Else
            If pPID > "" Then
                cmd.CommandText = "vGetPOLSOS"
            Else
                cmd.CommandText = "vGetPOLSOS2"
            End If
        End If
        Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
        cmd.Parameters.Append prm
        Set prm = cmd.CreateParameter("@COTPID", adInteger, adParamInput, , pCOTPID)
        cmd.Parameters.Append prm
        Set prm = cmd.CreateParameter("@STAFFID", adInteger, adParamInput)
        prm.Value = pSTAFFID
        cmd.Parameters.Append prm
        If pETABefore < CDate("1990-1-1") Then
            Set prm = cmd.CreateParameter("@PRIORTO", adDate, adParamInput, , ReverseDate(CDate("2200-01-01")))
            cmd.Parameters.Append prm
        Else
            Set prm = cmd.CreateParameter("@PRIORTO", adDate, adParamInput, , ReverseDate(pETABefore))
            cmd.Parameters.Append prm
        End If
        If pPID > "" Then
            Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput)
            prm.Value = pPID
            cmd.Parameters.Append prm
        End If
    End If
    
    
    Set prm = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set rs = cmd.execute()
    
    Set cmd = Nothing
    
    Set objPB = New PropertyBag
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    On Err GoTo H
    Set rs2 = New ADODB.Recordset
    i = 0

    Do While Not rs.eof 'And i <= pQTYtoReturn
        With udtProps
            If Not pSkipactions Then
            strAction = ""
            rs2.open "SELECT * FROM tPOLA WHERE POLA_POL_ID = " & FNN(rs!POL_ID) & " ORDER BY POLA_ID DESC", oPC.COShort, adOpenForwardOnly
            Do While Not rs2.eof
                strAction = strAction & IIf(Len(strAction) > 0, vbCrLf, "") & Format(rs2!POLA_Date, "dd/mm/yyyy") & ": " & ConvertPOLActionCodes(rs2!POLA_Code, bTmp) & ": " & FNS(rs2!POLA_Report)
                If strAction = ": ," Then strAction = ""
                rs2.MoveNext
            Loop
            rs2.Close
            Else
                strAction = ""
            End If
            udtProps.Actions = strAction
            udtProps.LastAction = FNS(rs!POL_LastAction)
            udtProps.LastActionDate = FND(rs!POL_LastActionDate)
            udtProps.ETA = FND(rs!POL_ETA)
            udtProps.Ref = FNS(rs!POL_REF)
            udtProps.Price = FNN(rs!POL_Price)
            udtProps.ForeignPrice = FNN(rs!POL_ForeignPrice)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyReceivedSoFar = FNN(rs!POL_QTYReceivedSoFar)
            udtProps.Discount = FNDBL(rs!POL_Discount)
            udtProps.POLID = FNN(rs!POL_ID)
            udtProps.PID = FNS(rs!POL_P_ID)
            udtProps.DOCCode = FNS(rs!TRCODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CFormat = FNS(rs!CFormat)
            udtProps.CDivisor = FNN(rs!CDivisor)
            udtProps.CForeignDivisor = FNN(rs!CDivisor)
            udtProps.CFactor = FNDBL(rs!CFactor)
            udtProps.COLID = FNN(rs!COLID)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyOutstanding = udtProps.QtySS + udtProps.QtyFirm - FNN(rs!POL_QTYReceivedSoFar)
            udtProps.QtyFirmOutstanding = udtProps.QtyFirm - FNN(rs!QTYFirmReceived)
            udtProps.QtySSOutstanding = udtProps.QtySS - FNN(rs!QTYSSReceived)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.MainAuthor = FNS(rs!P_MainAuthor)
            udtProps.code = FNS(rs!P_Code)
            udtProps.EAN = FNS(rs!P_EAN)
            udtProps.CodeF = FNS(rs!CodeF)
            udtProps.Supplier = FNS(rs!TP_Name)
            udtProps.SupplierID = FNS(rs!TPID)
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
'    pEOF = rs.EOF
    Set rs2 = Nothing
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    If Err = 7 Then
        MsgBox "There are too many records to fetch. Select fewer records by selecting a period further back."
        GoTo EXIT_Handler
    End If
    GoTo EXIT_Handler
    Resume

    Exit Function
errHandler:
    ErrPreserve
    If Err = 7 Then
        MsgBox "There are too many records to fetch. Select fewer records by selecting a period further back."
        GoTo EXIT_Handler
    End If

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLsOS_P.Fetch(pETABefore,pTPID,pPID,pStaffID)", Array(pETABefore, pTPID, pPID, _
         pSTAFFID)
End Function


Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set rs = New ADODB.Recordset

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLsOS_P.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
  '  rs.Close
    Set rs = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLsOS_P.Class_Terminate"
End Sub
