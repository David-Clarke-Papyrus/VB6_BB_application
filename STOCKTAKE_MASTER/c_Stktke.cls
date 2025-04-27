VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_stktke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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

Public Function Item(ByVal Index As Variant) As d_Stktke
  Set Item = colDisplay(Index)
End Function
Public Sub Load()
  ' load data from database
  Fetch
End Sub

Private Sub Fetch()
  Dim rs As New ADODB.Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim objDisplay As d_Stktke
  
  strSQL = "SELECT tblTA.TR_Printed, tblTA.TR_Void, tblTA.tblTA_ID, " _
            & "tblTA.TR_CaptureDate, tblTA.TR_Date, tblTA.TR_Code, " _
            & "tblTA.TR_SP, StockAdjustment.SA_SP_ID, Contact.CNT_Name " _
            & "FROM (tblTA INNER JOIN StockAdjustment " _
            & "ON tblTA.Transaction_ID = StockAdjustment.SA_TR_ID) " _
            & "INNER JOIN Contact ON StockAdjustment.SA_SP_ID = Contact.Contact_ID;"
            
  rs.Open strSQL, oADODBConn, adOpenDynamic, adLockReadOnly
  
  Do While Not rs.EOF
    Set objDisplay = New d_Stktke
    With objDisplay
        .ID = rs("Transaction_ID")
        .Code = rs("TR_Code")
        .CaptureDate = rs("TR_CaptureDate")
        .NominalDate = rs("TR_Date")
        .OperatorName = rs("CNT_Name")
        colDisplay.Add objDisplay
        Set objDisplay = Nothing
        rs.MoveNext
    End With
  Loop
  rs.Close
End Sub


