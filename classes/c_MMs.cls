VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_MM"
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

Public Function Item(ByVal Index As Variant) As d_MM
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function
Public Function Load(pPID As String)
    
    Fetch pPID
    
End Function
Public Function LoadAll(pPID As String, pDate As Date)
    
    FetchALL pPID, pDate
    
End Function

Private Sub Fetch(pPID As String)
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim oD As d_MM
Dim lngID As Long
Dim i As Long
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If oPC.IsMultiStore Then
        rs.open "SELECT zMM_IncVouchers.*,STORE_CODE FROM zMM_IncVouchers JOIN tPRODUCT ON PID = P_ID LEFT JOIN tPOSClient pc ON Station = StationName LEFT JOIN tStore s ON s.STORE_ID = pc.Store_ID WHERE PID = '" & pPID & "' ORDER BY TR_PROCESSINGDATE", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Else
        rs.open "SELECT * FROM zMM_IncVouchers JOIN tPRODUCT ON PID = P_ID WHERE PID = '" & pPID & "' ORDER BY TR_PROCESSINGDATE", oPC.COShort, adOpenForwardOnly, adLockReadOnly
      '  rs.open "SELECT * FROM zMM JOIN tPRODUCT ON PID = P_ID WHERE PID = '" & pPID & "' ORDER BY TR_PROCESSINGDATE", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    End If
    i = 1
    Do While Not rs.eof
        Set oD = New d_MM
            With oD
                .Qty = FNN(rs!Qty)
                .PID = FNS(rs!PID)
                .TRID = FNN(rs!TR_ID)
                .DOCCode = FNS(rs!TR_CODE)
                .SetDocdate FND(rs!TR_ProcessingDATE)
                .CaptureDate = FND(rs!TR_CaptureDate)
                .typ = FNS(rs!typ)
             ''   If oPC.IsMultiStore Then
                    .Station = FNS(rs!Station)
              ''  Else
                ''    .Station = ""
                ''End If
                .Seq = i
                colDisplay.Add oD
                Set oD = Nothing
            End With
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub FetchALL(pPID As String, pDate As Date)
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim oD As d_MM
Dim lngID As Long
Dim i As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM zMM_ALL JOIN tPRODUCT ON PID = P_ID WHERE PID = '" & pPID & "' AND TR_CaptureDATE >  {d '" & ReverseDate(pDate) & "'} ORDER BY TR_PROCESSINGDATE", oPC.COShort, adOpenForwardOnly, adLockReadOnly
  i = 1
    
    Do While Not rs.eof
        Set oD = New d_MM
            With oD
                .Qty = FNN(rs!Qty)
                .PID = FNS(rs!PID)
                .TRID = FNN(rs!TR_ID)
                .DOCCode = FNS(rs!TR_CODE)
                .SetDocdate FND(rs!TR_CaptureDate)
                .CaptureDate = FND(rs!TR_ProcessingDATE)
                .typ = FNS(rs!typ)
                .Seq = i
                
                colDisplay.Add oD
                Set oD = Nothing
            End With
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

