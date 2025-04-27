VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_BICCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As BICProps
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

Public Function Item(ByVal Index As Variant) As d_BICCode
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function Load()
Dim objPersist As c_BICCodes_P
    Set objPersist = New c_BICCodes_P
    SetSuperState objPersist.Fetch()
    Set objPersist = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_BICCode
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_BICCode
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp, oDisp.code
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub

Function FetchBICByCode(pBICCode As String) As d_BICCode
Dim oD As d_BICCode
    Set FetchBICByCode = Nothing
    For Each oD In colDisplay
        If oD.code = pBICCode Then
            Set FetchBICByCode = oD
            Exit For
        End If
    Next
On Error Resume Next
    Set FetchBICByCode = colDisplay(pBICCode)
End Function
Function FetchBICDescriptionsFromCodeSet(pIn As String) As String
Dim ar() As String
Dim strOut As String
Dim i As Integer
    ar = Split(pIn, ",")
    strOut = ""
    For i = 0 To UBound(ar)
        On Error Resume Next
        strOut = strOut & IIf(strOut > "", ",", "") & colDisplay(ar(i)).Description
    Next
    On Error GoTo 0
    FetchBICDescriptionsFromCodeSet = strOut
End Function

