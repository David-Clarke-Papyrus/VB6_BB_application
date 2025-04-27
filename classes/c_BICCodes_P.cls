VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_BICCodes_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch() As String
Dim lngCount As Long
Dim udtProps As BICProps
Dim udtData As BICData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tBIC"  ' WHERE TPID = " & TPID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        GoTo EXIT_Handler
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.ID = FNS(rs!BIC_ID)
            udtProps.code = FNS(rs!BIC_Code)
            udtProps.Description = FNS(rs!BIC_Description)
            udtProps.Level = FNS(rs!BIC_Level)
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




