VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_IG_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Fetch(pTPID As Long) As String
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As IGProps
Dim udtData As IGData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
Dim OpenResult As Integer

    strSQL = "SELECT TPIG_TP_ID, TPIG_IG_ID FROM tTP_IG WHERE TPIG_TP_ID =" & pTPID
    Set objPB = New PropertyBag
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenDynamic, adLockReadOnly
  
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
            .IGID = rs!TPIG_IG_ID
            .TPID = rs!TPIG_TP_ID
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
End Function

