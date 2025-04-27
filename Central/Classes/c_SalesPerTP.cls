VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_SalesPerTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mudtProps As DocsTPProps
Private colDisplay As Collection
Dim lngTPID As Long

Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_SalesPerTP
  Set Item = colDisplay(Index)
End Function

Public Function Load(pTPID As Long, Optional pRecentOnly As Boolean)
Dim oPersist As c_SalesPerTP_P
    lngTPID = pTPID
    Set oPersist = New c_SalesPerTP_P
    SetSuperState oPersist.Fetch(pTPID, 0, pRecentOnly)
    Set oPersist = Nothing
End Function
Public Function Reload(Optional pType As Integer)
Dim oPersist As c_SalesPerTP_P
    Set oPersist = New c_SalesPerTP_P
    SetSuperState oPersist.Fetch(lngTPID, pType)
    Set oPersist = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_SalesPerTP
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_SalesPerTP
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


