VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Operations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim FLastDayendDate As Date
Dim FLastCompactDate As Date

Private colDisplay As Collection

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set colDisplay = New Collection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.Class_Initialize"
    
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = colDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.Count"
    
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = colDisplay.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.NewEnum"
    
End Function

Public Function Item(ByVal Index As Variant) As d_operation
    On Error GoTo errHandler
  Set Item = colDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.Item(Index)", Index, EA_NORERAISE
    
End Function
Public Property Get LastUpdateDate() As Date
    LastUpdateDate = FLastDayendDate
End Property
Public Property Get LastCompactDate() As Date
    LastCompactDate = FLastCompactDate
End Property
Public Property Get LastUpdateDateFormatted() As String
    On Error GoTo errHandler
    LastUpdateDateFormatted = Format(FLastDayendDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.LastUpdateDateFormatted"
    
End Property
Public Property Get LastCompactDateFormatted() As String
    On Error GoTo errHandler
    LastCompactDateFormatted = Format(FLastCompactDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.LastCompactDateFormatted"
    
End Property
Public Sub Load(Optional pType As OpTypes)
    On Error GoTo errHandler
    If Not IsMissing(pType) Then
        Fetch pType
    Else
        Fetch
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.Load"
End Sub

Private Sub Fetch(Optional pType As OpTypes)
    On Error GoTo errHandler
  Dim rs As New ADODB.Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim objDisplay As d_operation
    If Not IsMissing(pType) Then
        If pType = enExportGroup Then
            strWHERE = "OP_Type = 9 or OP_TYPE = 11 or OP_Type = 12"
            rs.Open "SELECT TOP 300 * FROM tOperation WHERE " & strWHERE & " ORDER BY OP_StartedAT", oPC.COShort, adOpenDynamic, adLockReadOnly
        Else
            rs.Open "SELECT TOP 300 * FROM tOperation WHERE OP_TYPE = " & pType & " ORDER BY OP_StartedAT", oPC.COShort, adOpenDynamic, adLockReadOnly
        End If
    Else
        rs.Open "SELECT TOP 300 * FROM tOperation ORDER BY OP_StartedAT", oPC.COShort, adOpenDynamic, adLockReadOnly
    End If
  
  Do While Not rs.EOF
    Set objDisplay = New d_operation
    With objDisplay
        .ID = rs("OP_ID")
        Select Case rs.Fields("OP_Type")
            Case NielsenSales: .TypeName = "Sales export - Nielsen"
            Case StockSharing: .TypeName = "Stock sharing export"
            Case LoyaltyScheme: .TypeName = "Loyalty export"
        End Select
        .StartedAt = FND(rs("OP_StartedAT"))
        .Endedat = FND(rs("OP_EndedAt"))
        .NominalDate = FND(rs("OP_NominalDate"))
        .Result = FNN(rs("OP_Result"))
        colDisplay.Add objDisplay
        Set objDisplay = Nothing
        rs.MoveNext
    End With
  Loop
  rs.Close
'ERRH:
'MsgBox Error
'Exit Sub
'Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Operations.Fetch"
End Sub

