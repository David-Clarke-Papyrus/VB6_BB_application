VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLsperSUPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private colDisplay As Collection
Private colFilterSet As Collection
Dim bFilterOn As Boolean

Private Sub Class_Initialize()
  Set colDisplay = New Collection
  Set colFilterSet = New Collection
End Sub
Private Sub Class_Terminate()
  Set colDisplay = Nothing
  Set colFilterSet = Nothing
End Sub

Public Function Count() As Long
Attribute Count.VB_MemberFlags = "40"
  Count = colDisplay.Count
End Function
Public Function CountFilter() As Long
  CountFilter = colFilterSet.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If bFilterOn Then
        Set NewEnum = colFilterSet.[_NewEnum]
    Else
        Set NewEnum = colDisplay.[_NewEnum]
    End If
End Function

Public Function Item(ByVal Index As Variant) As d_POLine
  Set Item = colDisplay(Index)
End Function
Public Function FilterItem(ByVal Index As Variant) As d_POLine
  Set FilterItem = colFilterSet(Index)
End Function
Public Function ClearFilter()
    bFilterOn = False
End Function

Public Function Load(pTPID As Long)
Dim objPersist As c_POLsPerSUPP_P
    Set objPersist = New c_POLsPerSUPP_P
        SetSuperState objPersist.Fetch(pTPID)
    Set objPersist = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim oDisp As d_POLine
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_POLine
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp, oDisp.POLID & "k"
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


Public Function FilterByPID(pPID As String, pNewRow As Boolean)
Dim oPOL As d_POLine
    Set colFilterSet = Nothing
    Set colFilterSet = New Collection
    For Each oPOL In colDisplay
        If oPOL.PID = pPID And (oPOL.Applied = False Or pNewRow = False) Then
            colFilterSet.Add oPOL
        End If
    Next
  bFilterOn = True
End Function

Public Function FindByPOLID(pPOLID As Long) As d_POLine
Dim oPOL As d_POLine
    For Each oPOL In colDisplay
        If oPOL.POLID = pPOLID Then
            Set FindByPOLID = oPOL
        End If
    Next
End Function
