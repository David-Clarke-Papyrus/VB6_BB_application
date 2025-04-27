VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_OSAPPRO"
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
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_OSAPPRO
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function Load(pPID As String)
    
    Fetch pPID
    
End Function

Private Sub Fetch(pPID As String)
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim oD As d_OSAPPRO
Dim lngID As Long

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If oPC.ShowAllAppros = True Then
        rs.open "SELECT * FROM vALLAPPs WHERE APPL_P_ID = '" & pPID & "'", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Else
        rs.open "SELECT * FROM vOSAPPs WHERE APPL_P_ID = '" & pPID & "'", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    End If
    
    Do While Not rs.eof
        Set oD = New d_OSAPPRO
            With oD
                .DOCCode = FNS(rs!TR_CODE)
                .DOCDate = FND(rs!TR_DATE)
                .APPQty = FNN(rs!APPL_Qty)
                .APPReturned = FNN(rs!APPL_QtyReturned)
                .TRID = FNN(rs!TR_ID)
                .TPNAME = FNS(rs!TP_Name)
                .CaptureDate = FND(rs!TR_CaptureDate)
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

