VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLS"
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

Public Function Item(ByVal Index As Variant) As d_COL
  Set Item = colDisplay(Index)
End Function

Public Sub Load(pBranchCode As String, dteFrom As Date, dteTo As Date)
Dim oPersist As c_COLS_P

    Set oPersist = New c_COLS_P
    SetSuperState oPersist.Fetch(pBranchCode, dteFrom, dteTo)
    Set oPersist = Nothing

End Sub
Public Sub LoadRecordset(pBranchCode As String, dteFrom As Date, dteTo As Date, rs As ADODB.Recordset)
Dim oPersist As c_COLS_P

    Set oPersist = New c_COLS_P
    oPersist.FetchRecordset pBranchCode, dteFrom, dteTo, rs
    Set oPersist = Nothing

End Sub

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oCOL As d_COL
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oCOL = New d_COL
        oCOL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oCOL
        Set oCOL = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub





