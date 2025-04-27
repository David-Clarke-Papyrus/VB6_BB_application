VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_ZSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As ZSessionProps
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

Public Function Item(ByVal Index As Variant) As d_ZSession
  Set Item = colDisplay(Index)
End Function

Public Function Load(pSince As Date, pExactdate As Date, pExchangeNumber As Long)
Dim oPersist As c_ZSession_P
    Set oPersist = New c_ZSession_P
    SetSuperState oPersist.Fetch(pSince, pExactdate, pExchangeNumber)
    Set oPersist = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_ZSession
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_ZSession
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


Private Sub Class_Terminate()
    Set colDisplay = Nothing
End Sub
