VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CATCHK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Function Item(ByVal Index As Variant) As d_CATCHK
  Set Item = colDisplay(Index)
End Function

Public Sub Load(ByRef pblnNoRecsReturned As Boolean, Optional CATCode As String, Optional CATShortCode As String, Optional pCode As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pSupplierID As Long, Optional pPEAN As String, Optional pPID As String)
Dim objPersist As c_CatChk_P

    Set objPersist = New c_CatChk_P
    SetSuperState objPersist.Fetch(pblnNoRecsReturned, CATCode, CATShortCode, pCode, pdteDate1, pdteDate2, , pPEAN, pPID)
    Set objPersist = Nothing

End Sub

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim odCATCHK As d_CATCHK
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odCATCHK = New d_CATCHK
        odCATCHK.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add odCATCHK
        Set odCATCHK = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



