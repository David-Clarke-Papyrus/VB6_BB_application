VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Cashups"
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

Public Function Item(ByVal Index As Variant) As d_Cashup
  Set Item = colDisplay(Index)
End Function

Public Sub Load(pBranchCode As String, pDiscrepancy As String, pTIllpoint As String, dteFrom As Date, dteTo As Date, pStatus As String)
Dim oPersist As c_Cashups_P

    Set oPersist = New c_Cashups_P
    SetSuperState oPersist.Fetch(pBranchCode, pDiscrepancy, pTIllpoint, dteFrom, dteTo, pStatus)
    Set oPersist = Nothing

End Sub
Public Sub LoadRecordset(pBranchCode As String, pDiscrepancy As String, pTIllpoint As String, dteFrom As Date, dteTo As Date, pStatus As String, rs As ADODB.Recordset)
Dim oPersist As c_Cashups_P

    Set oPersist = New c_Cashups_P
    oPersist.FetchRecordset pBranchCode, pDiscrepancy, pTIllpoint, dteFrom, dteTo, pStatus, rs
    Set oPersist = Nothing

End Sub

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oCashup As d_Cashup
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oCashup = New d_Cashup
        oCashup.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oCashup
        Set oCashup = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub




