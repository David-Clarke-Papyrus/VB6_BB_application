VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLsPerCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As COLProps
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

Public Function Item(ByVal Index As Variant) As d_COLine
  Set Item = colDisplay(Index)
End Function
Public Function FilterItem(ByVal Index As Variant) As d_COLine
  Set FilterItem = colFilterSet(Index)
End Function
Public Function ClearFilter()
    bFilterOn = False
End Function

Public Function Load(pTPID As Long, pPID As String, Optional pCOLID As Long)
    On Error GoTo errHandler
'Dim objPersist As c_COLsPerCust_P
'    Set objPersist = New c_COLsPerCust_P
'        SetSuperState objPersist.Fetch(pTPID)
Dim objPersist As c_COLsOD_P
    Set objPersist = New c_COLsOD_P
        SetSuperState objPersist.Fetch(, pTPID, pPID, , pCOLID)
    Set objPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_COLsPerCust.Load(pTPID,pPID,pCOLID)", Array(pTPID, pPID, pCOLID)
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


Public Function FilterByPID(pPID As String)
Dim oCOL As d_COLine
    Set colFilterSet = Nothing
    Set colFilterSet = New Collection
    For Each oCOL In colDisplay
        If oCOL.PID = pPID Then
            colFilterSet.Add oCOL
        End If
    Next
  bFilterOn = True
End Function

