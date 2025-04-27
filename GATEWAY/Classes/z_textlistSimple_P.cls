VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_textlistSimple_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum ListType
    ltCNote = 1
    ltEntity = 2
    ltGrower = 3
End Enum

Public Function Fetch(ListName As ListType, Optional pParam As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim objPB As PropertyBag
Dim udtProps As TextListSProps
Dim udtData As TextListSData
Dim lngCount As Long
Dim strSQL As String

    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag

    Select Case ListName
        Case 1
                strSQL = "SELECT CNote FROM vUniqueCNotes WHERE ENTITY = '" & pParam & "' ORDER BY CNote"
        Case 2
            If pParam > "" Then
                strSQL = "SELECT Entity FROM vUniqueEntities WHERE Entity = '" & pParam & "' ORDER BY Entity"
            Else
                strSQL = "SELECT Entity FROM vUniqueEntities ORDER BY Entity"
            End If
        Case 3
            strSQL = "SELECT * FROM vGrowers"
    End Select
    rs.Open strSQL, oPC.COShort, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        If IsNull(rs(0)) Then GoTo Tail
        udtProps.Item = rs(0)
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
'ERRH:
''  rs.Close
'  Set rs = Nothing
'
'
'  Err.Raise Err.Number
'  Exit Function
'  Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_textlistSimple_P.Fetch(ListName,pParam)", Array(ListName, pParam), EA_NORERAISE
    
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

