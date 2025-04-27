VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_OSSOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colDisplay As Collection

Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_OSSOrder
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function
Public Function Load(PID As String)
    
    Fetch PID
    
End Function

Private Sub Fetch(pPID As String)

Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim oD As d_OSSOrder
Dim lngID As Long

If Not pPID > "" Then Exit Sub
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM vOSPOs WHERE POL_P_ID = '" & pPID & "'", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.eof
        Set oD = New d_OSSOrder
            With oD
                .Firm = FNN(rs!POL_QtyFirm)
                .TRID = FNN(rs!TR_ID)
                .SS = FNN(rs!POL_QtySS)
                .QtyReceived = FNN(rs!POL_QTYReceivedSoFar)
                .DOCCode = FNS(rs!TR_CODE)
                .DOCDate = FND(rs!TR_DATE)
                .CaptureDate = FND(rs!TR_CaptureDate)
                .ETA = FND(rs!POL_ETA)
                .Name = FNS(rs!TP_Name)
                colDisplay.Add oD
                Set oD = Nothing
            End With
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub
