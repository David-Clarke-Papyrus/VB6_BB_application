VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_OSCOrder"
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

Public Function Item(ByVal Index As Variant) As d_OSCOrder
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function Load(PID As String)
    
    Fetch PID
    
End Function

Private Sub Fetch(pPID As String)
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim oD As d_OSCOrder
Dim lngID As Long

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM vOSCOs WHERE COL_P_ID = '" & pPID & "'", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.eof
        Set oD = New d_OSCOrder
            With oD
                .DOCCode = FNS(rs!TR_CODE)
                If IsNull(rs!TR_DATE) Then
                    .DOCDate = CDate(0)
                Else
                    .DOCDate = FND(rs!TR_DATE)
                End If
                .TPNAME = FNS(rs!TP_Name)
                .COLQty = FNN(rs!COL_Qty)
                .COLCollected = FNN(rs!COL_QtyDispatched)
                .Ref = FNS(rs!COL_Ref)
               ' .Email = FNS(rs!ADD_Email)
                .TRID = FNN(rs!TR_ID)
                .CaptureDate = FND(rs!TR_CaptureDate)
                .COLID = FNN(rs!COL_ID)
                .ETA = FND(rs!COL_ETA)

                colDisplay.Add oD
                Set oD = Nothing
                rs.MoveNext
            End With
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub
