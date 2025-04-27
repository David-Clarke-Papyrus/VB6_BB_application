VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_OrdersPerSupplier_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(TPID As Long, Optional pDLID As Long) As String
Dim lngCount As Long
Dim udtProps As OPSProps
Dim udtData As OPSData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If pDLID > 0 Then
        strSQL = "SELECT * FROM vOrdersPerSupplier WHERE TR_TP_ID = " & TPID & " AND POL_DL_ID = " & pDLID
    Else
        strSQL = "SELECT * FROM vOrdersPerSupplier WHERE TR_TP_ID = " & TPID
    End If
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly
    If rs.EOF Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.EOF
        With udtProps
            .dateOfOrder = FND(rs!TR_DATE)
            .Price = FNN(rs!POL_Price)
            udtProps.Code = FNS(rs!P_Code)
            .Title = FNS(rs!P_Title)
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

