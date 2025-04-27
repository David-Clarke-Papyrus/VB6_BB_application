VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_JNL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As dJNLProps

Private colDisplay As Collection

Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_JNL
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function




Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pTPID As Long, Optional pACNO As String, Optional pRef As String, Optional pDate1 As Date, _
                Optional pDate2 As Date)
Dim objPersist As c_JNL_P
    Set objPersist = New c_JNL_P
    SetSuperState objPersist.Fetch(pNoRecsReturned, pTPID, pACNO, pRef, pDate1, pDate2)
    Set objPersist = Nothing
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oD As d_JNL
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oD = New d_JNL
        oD.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oD
        Set oD = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



