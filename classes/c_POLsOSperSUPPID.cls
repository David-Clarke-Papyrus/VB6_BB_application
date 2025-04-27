VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLsOSperSUPPID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mudtProps As dCOLProps
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

Public Function Item(ByVal Index As Variant) As d_COLine
  Set Item = colDisplay(Index)
End Function

Public Function Load(pPID As String, Optional pOrderType As enumCOType)
Dim objPersist As c_COLsPerPID_P
    Set objPersist = New c_COLsPerPID_P
 '   If pOrderType > 0 Then
        SetSuperState objPersist.Fetch(pPID, enWant)
 '   Else
 '       SetSuperState objPersist.Fetch(pPID)
 '   End If
    Set objPersist = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim oDisp As d_COLine
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_COLine
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp, oDisp.COLID & "k"
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



