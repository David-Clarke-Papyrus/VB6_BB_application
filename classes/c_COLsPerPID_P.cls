VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLsPerPID_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pPID As String, Optional pOrderType As enumCOType) As String
Dim lngCount As Long
Dim udtProps As dCOLProps
Dim udtData As dCOLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim objDisplay As d_COLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    If pOrderType > 0 Then
        strSQL = "SELECT * FROM vWantsPerPID WHERE COL_P_ID = '" & pPID & "' AND CO_ORDERTYPE = " & pOrderType
    Else
        strSQL = "SELECT * FROM COLsPerPID WHERE COL_P_ID = '" & pPID & "'"
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
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.COLID = FNN(rs!COL_ID)
            udtProps.Qty = FNN(rs!COL_Qty)
            udtProps.Note = FNS(rs!COL_Note)
            udtProps.Ref = FNS(rs!COL_Ref)
            udtProps.DOCDate = FND(rs!TR_DATE)
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.CustInitials = FNS(rs!TP_Initials)
            udtProps.CustTitle = FNS(rs!TP_Title)
            udtProps.CustName = FNS(rs!TP_Name)
            udtProps.WantDate = FND(rs!COL_WantDate)
            udtProps.TRID = FNN(rs!TR_ID)
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
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function


