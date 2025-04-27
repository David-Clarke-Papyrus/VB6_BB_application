VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Function Item(ByVal Index As Variant) As d_COSR
  Set Item = colDisplay(Index)
End Function

Public Function LoadByProd(PID As String)
Dim objPersist As c_COSR_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_COSR_P
    SetSuperState objPersist.Fetch(bRecordsReturned, , , , , , PID)
    Set objPersist = Nothing
End Function


Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pTPID As Long, Optional pACNO As String, Optional pRef As String, Optional pDate1 As Date, _
                Optional pDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pPIID As Long, Optional pPF As String)
Dim oCOSR As c_COSR_P
'        load data from database
    Set oCOSR = New c_COSR_P
    SetSuperState oCOSR.Fetch(pNoRecsReturned, pTPID, pACNO, pRef, pDate1, pDate2, pPID, pPCode)
    Set oCOSR = Nothing
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oD As d_COSR
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oD = New d_COSR
        oD.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oD
        Set oD = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



