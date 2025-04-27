VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_GDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dim mudtProps As dInvoiceProps

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

Public Function Item(ByVal Index As Variant) As d_GDN
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function LoadByProd(PID As String)
Dim objPersist As c_GDN_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_GDN_P
    SetSuperState objPersist.Fetch(bRecordsReturned, , , , , , , PID)
    Set objPersist = Nothing
End Function


Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pTPID As Long, Optional pACNO As String, Optional pRef As String, Optional pDocRef As String, Optional pDate1 As Date, _
                Optional pDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pPIID As Long, _
                Optional pUseProcessingDate As Boolean, Optional pUNissued As Boolean, Optional AllGDNs As Boolean)
Dim objInvoicesPersist As c_GDN_P
'        load data from database
    Set objInvoicesPersist = New c_GDN_P
    SetSuperState objInvoicesPersist.Fetch(pNoRecsReturned, pTPID, pACNO, pRef, pDocRef, pDate1, pDate2, pPID, pPCode, pPIID, pUseProcessingDate, pUNissued, AllGDNs)
    Set objInvoicesPersist = Nothing
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oD As d_GDN
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oD = New d_GDN
        oD.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oD
        Set oD = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



