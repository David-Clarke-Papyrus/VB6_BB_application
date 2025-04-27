VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_textlist_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum ListType
    ltCNote = 1
    ltEntity = 2
End Enum

Public Function Fetch(ListName As ListType, Optional pParam As String) As String
Dim rs As ADODB.Recordset
Dim objPB As PropertyBag
Dim udtProps As TextListProps
Dim udtData As TextListData
Dim lngCount As Long
Dim strSQL As String

    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag
    On Error GoTo ERRH

    Select Case ListName
        Case 1
                strSQL = "SELECT CNote FROM cUniqueCNotes ORDER BY CNote"
        Case 2
                strSQL = "SELECT Entity FROM vUniqueEntities ORDER BY Entity"
    End Select
    rs.Open strSQL, oPC.CO, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        If IsNull(rs(1)) Then GoTo Tail
        udtProps.Item = rs(1)
        udtProps.Key = rs(0)
        If rs.Fields.Count > 2 Then
            udtProps.F3 = FNS(rs(2))
        Else
            udtProps.F3 = ""
        End If
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
Tail:
        rs.MoveNext
    Loop
    rs.Close
    objPB.WriteProperty "Count", lngCount
    Fetch = objPB.Contents

    Set rs = Nothing
    Exit Function
ERRH:
    On Error Resume Next
  rs.Close
  Set rs = Nothing
  
  
  Err.Raise Err.Number
  Exit Function
  Resume
End Function

'#If conWin2000 Then
'
'Private Sub ObjectControl_Activate()
'  Set mobjContext = GetObjectContext
'End Sub
'
'Private Function ObjectControl_CanBePooled() As Boolean
'  ObjectControl_CanBePooled = False
'End Function
'
'Private Sub ObjectControl_Deactivate()
'  Set mobjContext = Nothing
'End Sub
'#EndIF

