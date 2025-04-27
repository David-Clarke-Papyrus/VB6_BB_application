VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mudtProps As CSLProps

Private colDisplay As Collection

Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_CSL
  Set Item = colDisplay(Index)
End Function

'Public Function DisplayInvoice() As d_Order
'    Set DisplayInvoice = New d_Order
'End Function

'Public Function LoadByProd(pTRID As Long)
'Dim objPersist As c_CSL_P
'Dim bRecordsReturned As Boolean
'    Set objPersist = New c_CSL_P
'    SetSuperState objPersist.Fetch(pTRID)
'    Set objPersist = Nothing
'End Function


Public Sub Load(pTRID As Long)
Dim oPersist As c_CSL_P
'        load data from database
    Set oPersist = New c_CSL_P
    SetSuperState oPersist.Fetch(pTRID)
    Set oPersist = Nothing
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim odSO As d_CSL
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odSO = New d_CSL
        odSO.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add odSO
        Set odSO = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



