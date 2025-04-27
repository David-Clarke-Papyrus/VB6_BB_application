VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_OrdersPerSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mudtProps As OPSProps
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

Public Function Item(ByVal Index As Variant) As d_OrdersPerSupplier
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function Load(pTPID As Long, Optional pDLID As Long)
Dim objOPSPersist As c_OrdersPerSupplier_P
    Set objOPSPersist = New c_OrdersPerSupplier_P
    SetSuperState objOPSPersist.Fetch(pTPID, pDLID)
    Set objOPSPersist = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_OrdersPerSupplier
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_OrdersPerSupplier
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub

