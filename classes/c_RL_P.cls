VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_RL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pTRID As Long) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dRLProps
Dim udtData As dRLData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim bIncludesLineDetails As Boolean
Dim OpenResult As Integer
On Error GoTo H
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oBatch.RunGetRecordset("Select tRL.*,DEL_SUPPLIERINVOICEDATE,dbo.PubCode(P_Code) as Pubcode,dbo.CodeF(P_Code,P_EAN,1) as codeF,P_ID,P_Title,P_MainAuthor,P_EAN,P_Section,P_Code,P_QtyOnHand FROM tRL JOIN tProduct ON P_ID = RL_P_ID LEFT JOIN tDELL ON RL_DELL_ID = DELL_ID Left JOIN tDEL ON DELL_TR_ID = DEL_ID WHERE RL_TR_ID = " & pTRID & " ORDER BY P_Title", enText, Array(), "", rs)
   ' bIncludesLineDetails = True
    
    Set objPB = New PropertyBag
    
    If rs.eof Then
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = FNN(rs!RL_TR_ID)
            udtProps.RLID = FNN(rs!RL_ID)
            udtProps.ProductDescription = FNS(rs!P_Title) & " / " & FNS(rs!P_MainAuthor)
            udtProps.PID = FNS(rs!P_ID)
            udtProps.code = FNS(rs!CodeF)
            udtProps.Section = FNS(rs!P_Section)
            udtProps.QtySystem = FNN(rs!P_QtyOnHand)
            udtProps.QtyRequested = FNN(rs!RL_QtyRequested)
            udtProps.QtyApproved = FNN(rs!RL_QtyApproved)
            udtProps.QtyReturned = FNN(rs!RL_QtyReturned)
            udtProps.QtyCounted = FNN(rs!RL_QtyCounted)
            udtProps.val = FNN(rs!RL_Price) * FNN(rs!RL_QtyReturned) * (((100 - FNN(rs!RL_Discount)) / 100))
            udtProps.SupplierInvoiceRef = FNS(rs!RL_SiREF)
            udtProps.SupplierInvoiceDate = FND(rs!DEL_SUpplierInvoiceDate)
            udtProps.Note = FNS(rs!RL_Memo)
            udtProps.Pubcode = rs!Pubcode
            udtProps.EAN = FNS(rs!P_EAN)
            udtProps.DELLID = FNN(rs!RL_DELL_ID)
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
    Set oBatch = Nothing
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function




