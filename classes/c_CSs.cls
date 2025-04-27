VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CSs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As CSProps

Private colDisplay As Collection

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set colDisplay = New Collection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = colDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = colDisplay.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As d_CS
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = colDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.Item(Index)", Index
End Function

'Public Function DisplayInvoice() As d_Invoice
'    Set DisplayInvoice = New d_Invoice
'End Function
Public Function FindByCSLID(pCSLID As Long)
    On Error GoTo errHandler
Dim oD As Variant

    For Each oD In colDisplay
        If oD.CSLID = pCSLID Then
            Set FindByCSLID = oD
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.FindByCSLID(pCSLID)", pCSLID
End Function
Public Function LoadByProd(PID As String)
    On Error GoTo errHandler
Dim objPersist As c_CSs_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_CSs_P
    SetSuperState objPersist.Fetch(bRecordsReturned, , , , PID)
    Set objPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.LoadByProd(PID)", PID
End Function
Public Function LoadByZID(ZID As String)
    On Error GoTo errHandler
Dim objPersist As c_CSs_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_CSs_P
    SetSuperState objPersist.FetchByZ_ID(bRecordsReturned, ZID)
    Set objPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.LoadByZID(ZID)", ZID
End Function


Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pDocCOde As String, Optional pDate1 As Date, Optional pDate2 As Date, Optional pPID As String, Optional pCode As String)
    On Error GoTo errHandler
Dim objPersist As c_CSs_P
Dim strResult As String

    Set objPersist = New c_CSs_P
        SetSuperState objPersist.Fetch(pNoRecsReturned, pDocCOde, pDate1, pDate2, pPID, pCode)
    Set objPersist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.Load(pNoRecsReturned,pDocCode,pDate1,pDate2,pPID,pCODE)", Array(pNoRecsReturned, _
         pDocCOde, pDate1, pDate2, pPID, pCode)
End Sub


Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oD As d_CS
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oD = New d_CS
        oD.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oD
        Set oD = Nothing
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CSs.SetSuperState(buffer)", buffer
End Sub

