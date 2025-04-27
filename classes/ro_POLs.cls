VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ro_POLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolItems As Collection
Private mParent As ro_PO

Dim oPOL As ro_POL
Attribute oPOL.VB_VarHelpID = -1


Private Sub Class_Initialize()
  Set mcolItems = New Collection
End Sub
Public Sub component(pParent As ro_PO)
    Set mParent = pParent
End Sub
Public Function Count() As Long
  Count = mcolItems.Count
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As ro_POL
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function



Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oPO As ro_POL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count
    For Each oPO In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oPO.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oPOL As ro_POL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oPOL = New ro_POL
      oPOL.SetParent mParent
      oPOL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      mcolItems.Add oPOL, oPOL.POLID & "k"
      Set oPOL = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub
Function FindLineByID(PID As Long) As ro_POL
Dim oPOL As ro_POL
    
    For Each oPOL In mcolItems
        If oPOL.POLID = PID Then
            Set FindLineByID = oPOL
            Exit For
        End If
    Next
End Function

Function CancelInactiveLines()
Dim oPOL As ro_POL
    
    For Each oPOL In mcolItems
        If oPOL.Fulfilled = "OS" Then
            If oPOL.QtyReceivedSoFar = 0 Then
                oPOL.CancelLine
            Else
                oPOL.MarkLineasFulfilled
            End If
        End If
    Next
End Function

