VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As RProps

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

Public Function Item(ByVal Index As Variant) As d_R
  Set Item = colDisplay(Index)
End Function
Public Property Get ReturnType() As String
    ReturnType = mudtProps.RType
End Property
'Public Function DisplayInvoice() As d_Order
'    Set DisplayInvoice = New d_Order
'End Function

Public Function LoadByProd(PID As String)
Dim objPersist As c_R_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_R_P
    SetSuperState objPersist.Fetch(bRecordsReturned, , , , , , PID)
    Set objPersist = Nothing
End Function


Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pTPID As Long, _
                Optional pACNO As String, _
                Optional pRef As String, _
                Optional pDate1 As Date, _
                Optional pDate2 As Date, _
                Optional pPID As String, _
                Optional pPCode As String, _
                Optional pUNissued As Boolean)
Dim oPersist As c_R_P
'        load data from database
    Set oPersist = New c_R_P
    SetSuperState oPersist.Fetch(pNoRecsReturned, pTPID, pACNO, pRef, pDate1, pDate2, , pPCode, pUNissued)
    Set oPersist = Nothing
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim odR As d_R
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odR = New d_R
        odR.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add odR, odR.TRID & "k"
        Set odR = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



