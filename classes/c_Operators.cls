VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Operators"
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

Public Function Item(ByVal Index As Variant) As d_operator
  Set Item = colDisplay(Index)
End Function

Public Sub Load(PID As Long)
  ' load data from database
  Fetch PID
End Sub

Private Sub Fetch(PID As Long)
  Dim rs As New ADODB.Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim objDisplay As d_operator
'  Dim oBatch As New z_Batch
'  oBatch.RunGetRecordset "Operators_qsel2", Query, Array(pID)
'  SELECT Contact.Contact_ID, Contact.CNT_Name, Contact.CNT_TP_ID
'From Contact, tblConfiguration
'WHERE (((Contact.CNT_TP_ID)=[CF_OwnerID]));
  rs.Open "Operators_qsel", oPC.CO, adOpenDynamic, adLockReadOnly
  
  Do While Not rs.EOF
    Set objDisplay = New d_operator
    With objDisplay
      .ID = rs("Contact_ID")
      .Fullname = rs("CNT_Name")
      
      colDisplay.Add objDisplay
      Set objDisplay = Nothing
      rs.MoveNext
    End With
  Loop
  rs.Close
End Sub


