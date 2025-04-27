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
Attribute VB_Exposed = False
Option Explicit
Dim FLastDayendDate As Date
Dim FLastCompactDate As Date

Private colDisplay As Collection

Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_operation
  Set Item = colDisplay(Index)
End Function
Public Property Get LastUpdateDate() As Date
    LastUpdateDate = FLastDayendDate
End Property
Public Property Get LastCompactDate() As Date
    LastCompactDate = FLastCompactDate
End Property
Public Property Get LastUpdateDateFormatted() As String
    LastUpdateDateFormatted = Format(FLastDayendDate, "dd/mm/yyyy")
End Property
Public Property Get LastCompactDateFormatted() As String
    LastCompactDateFormatted = Format(FLastCompactDate, "dd/mm/yyyy")
End Property
Public Sub Load()
  ' load data from database
  Fetch
End Sub

Private Sub Fetch()
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim strWHERE As String
Dim objDisplay As d_operation
Dim OpenResult As Integer

  On Error GoTo H
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
  rs.open "SELECT TOP 200 * FROM tOperation LEFT OUTER JOIN tStaffmember on OP_StartedByID = SM_ID ORDER BY OP_StartedAT DESC ", oPC.COShort, adOpenDynamic, adLockReadOnly
  
  Do While Not rs.eof
    Set objDisplay = New d_operation
    With objDisplay
        .ID = rs("OP_ID")
        Select Case rs.fields("OP_Type")
            Case Dayend: .TypeName = "Daily update"
            Case Compact: .TypeName = "Compact"
            Case Export: .TypeName = "Export"
            Case ClearOrders: .TypeName = "Deleting old orders"
            Case SupplierChange: .TypeName = "Supplier change"
            Case AllocationStockToOrders: .TypeName = "Order fulfilment"
            Case SimplifyTitles: .TypeName = "Strip 'The etc."
            Case ClearOldData: .TypeName = "Cleared old data"
            Case SupplierChange: .TypeName = "Changed supplier"
            Case NielsenSales: .TypeName = "Export sales data"
            Case BookDataWash: .TypeName = "Book data wash"
            Case LoyaltyScheme: .TypeName = "Export loyalty customer"
            Case StockSharing: .TypeName = "Export stock-sharing"
            Case PastelExportDebtors: .TypeName = "Pastel export debtors"
            Case PastelExportCreditors: .TypeName = "Pastel export creditors"
            Case PastelExportCustomers: .TypeName = "Pastel export customers"
            Case PastelImportCustomers: .TypeName = "Pastel import customers"
        End Select
        '.TypeName = rs("TypeName")
        .StartedAt = FND(rs("OP_StartedAT"))
        .Endedat = FND(rs("OP_EndedAt"))
        .NominalDate = FND(rs("OP_NominalDate"))
        .Result = FNN(rs("OP_Result"))
        .OperatorName = FNS(rs("SM_Name"))
        If .OperatorName = "" Then .OperatorName = "AUTO"
        If (FLastDayendDate < rs("OP_NominalDate")) And (rs.fields("OP_Type") = Dayend) And (rs.fields("OP_Result") = 1) Then
            FLastDayendDate = CDate(rs("OP_NominalDate"))
        End If
        If (FLastCompactDate < rs("OP_NominalDate")) And (rs.fields("OP_Type") = Compact) And (rs.fields("OP_Result") = 1) Then
            FLastCompactDate = CDate(rs("OP_NominalDate"))
        End If
        colDisplay.Add objDisplay
        Set objDisplay = Nothing
        rs.MoveNext
    End With
  Loop
  rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  Exit Sub
H:
MsgBox Error
Exit Sub
Resume
End Sub

