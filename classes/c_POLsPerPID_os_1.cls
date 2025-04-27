VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLsPerPID_os_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colDisplay As Collection
Private colFilterSet As Collection
Dim bFilterOn As Boolean
Event lngProgress(p As Long)
Event lngMax(p As Long)
Dim WithEvents objPersist As c_POLsPerPID_os_1_P
Attribute objPersist.VB_VarHelpID = -1
Dim iMaxRecords As Long
Dim iLastRecNum As Long
Dim iSetSize As Long
Dim bEOF As Boolean

Private Sub Class_Initialize()
  Set colDisplay = New Collection
  Set colFilterSet = New Collection
    Set objPersist = New c_POLsPerPID_os_1_P
End Sub
Private Sub Class_Terminate()
    Set colDisplay = Nothing
    Set colFilterSet = Nothing
    Set objPersist = Nothing
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
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
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function
Public Function FilterItem(ByVal Index As Variant) As d_POLine
  Set FilterItem = colFilterSet(Index)
End Function
Public Function ClearFilter()
    bFilterOn = False
End Function

Public Function Load(pPID As String)
        SetSuperState objPersist.Fetch(pPID)
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



