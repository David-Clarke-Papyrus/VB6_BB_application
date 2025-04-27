VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Error"
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

Public Function Item(ByVal Index As Variant) As d_Error
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
  Dim objDisplay As d_Error
  
  rs.Open "GetErrors", oPC.CO, adOpenDynamic, adLockReadOnly
  
  Do While Not rs.EOF
    Set objDisplay = New d_Error
    With objDisplay
      .ID = rs("Err_ID")
      .Description = rs("ERR_Description")
      .DateOfError = rs("ERR_Date")
      .FormName = rs("ERR_Form")
      .ReportName = rs("ERR_Report")
      .UserName = rs("ERR_User")
      .Number = rs("ERR_Number")
      
      colDisplay.Add objDisplay
      Set objDisplay = Nothing
      rs.MoveNext
    End With
  Loop
  rs.Close
End Sub


