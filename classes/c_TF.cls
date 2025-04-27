VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_TF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As TFProps

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

Public Function Item(ByVal Index As Variant) As d_TF
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

'Public Function DisplayInvoice() As d_Invoice
'    Set DisplayInvoice = New d_Invoice
'End Function
Public Function FindByTFLID(pTFLID As Long)
Dim oD As Variant

    For Each oD In colDisplay
        If oD.TFLID = pTFLID Then
            Set FindByTFLID = oD
            Exit For
        End If
    Next
End Function
Public Function LoadByProd(PID As String)
Dim objPersist As c_TF_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_TF_P
    SetSuperState objPersist.Fetch(bRecordsReturned, , , , PID)
    Set objPersist = Nothing
End Function


Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pDocCOde As String, Optional pDate1 As Date, Optional pDate2 As Date, _
Optional pPID As String, Optional pCode As String, Optional pUNissued As Boolean)
Dim objPersist As c_TF_P
Dim strResult As String

    Set objPersist = New c_TF_P
        SetSuperState objPersist.Fetch(pNoRecsReturned, pDocCOde, pDate1, pDate2, pPID, pCode, pUNissued)
    Set objPersist = Nothing
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oD As d_TF
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oD = New d_TF
        oD.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oD
        Set oD = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


