VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CSL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(pTRID As Long) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dCSLProps
Dim udtData As dCSLData
Dim objPB As PropertyBag
Dim lngCount As Long
    
    Set rs = New ADODB.Recordset
    
    rs.Open "Select * FROM vCSL WHERE CSL_TR_ID = " & pTRID, oPC.CO, adOpenKeyset, adLockOptimistic
    
    Set objPB = New PropertyBag
    
    If rs.EOF Then
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo ERRH
    Do While Not rs.EOF
        With udtProps
            udtProps.ID = rs!CSL_ID
            udtProps.Title = FNS(rs!P_Title)
            udtProps.Qty = FNS(rs!CSL_Qty)
            udtProps.DateTime = FND(rs!TR_Date)
            udtProps.EAN = FNN(rs!P_EAN)
            udtProps.Code = FNS(rs!P_Code)
            udtProps.Price = FNS(rs!CSL_Price)
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
    Set oBatch = Nothing
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function


