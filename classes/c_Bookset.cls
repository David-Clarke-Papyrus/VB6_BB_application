VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Bookset"
Attribute VB_GlobalNameSpace = True
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

Public Function Item(ByVal Index As Variant) As d_Bookset
  Set Item = colDisplay(Index)
End Function

Public Sub Load()
  ' load data from database
  Fetch
End Sub

Private Sub Fetch()
  Dim rs As New ADODB.Recordset
  Dim objDisplay As d_Bookset
  
  rs.Open "GetBooksets", oPC.CO, adOpenDynamic, adLockReadOnly
  
  Do While Not rs.EOF
    Set objDisplay = New d_Bookset
    With objDisplay
      .ProdSetID = rs("ProductSet_ID")
      .Description = rs("PS_Description")
      .ShortDescription = rs("PS_ShortDesc")
      
      colDisplay.Add objDisplay, .Description
      Set objDisplay = Nothing
      rs.MoveNext
    End With
  Loop
  rs.Close
End Sub


