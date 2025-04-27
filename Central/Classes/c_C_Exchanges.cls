VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Exchanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mudtProps As ExchangeProps
Private colDisplay As Collection
Private colCSLDisplay As Collection
Private colPayDisplay As Collection

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set colDisplay = New Collection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = colDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = colDisplay.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As d_Exchange
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = colDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.Item(Index)", Index
End Function

Public Function Load(dteFrom As Date, dteTo As Date, pExchangeNumber As Long, pBranchCode As String)
    On Error GoTo errHandler
Dim oPersist As c_Exchanges_P
    Set oPersist = New c_Exchanges_P
        SetSuperState oPersist.Fetch(dteFrom, dteTo, pExchangeNumber, pBranchCode)
    Set oPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.Load(dteFrom,dteTo)", Array(dteFrom, dteTo)
End Function
Public Function LoadCustomerRequests(pOSorALL As String)
    On Error GoTo errHandler
Dim oPersist As c_Exchanges_P
    Set oPersist = New c_Exchanges_P
    SetSuperState oPersist.FetchCustomerRequests(pOSorALL)
    Set oPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.LoadCustomerRequests(pOSorALL)", Array(pOSorALL)
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_Exchange
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_Exchange
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp, oDisp.ID
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.SetSuperState(buffer)", buffer
End Sub


Public Function LoadCSL(EXCHID As String)
    On Error GoTo errHandler
Dim oPersist As c_Exchanges_P
    Set oPersist = New c_Exchanges_P
    Set colCSLDisplay = Nothing
    Set colCSLDisplay = New Collection
        SetSuperStateCSL oPersist.FetchCSL(EXCHID)
    Set oPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.LoadCSL(EXCHID)", EXCHID
End Function
Friend Sub SetSuperStateCSL(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_CSL
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_CSL
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colCSLDisplay.Add oDisp
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.SetSuperStateCSL(buffer)", buffer
End Sub
Public Function CSLItem(ByVal Index As Variant) As d_CSL
    On Error GoTo errHandler
    Set CSLItem = colCSLDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.CSLItem(Index)", Index
End Function
Public Function CSLCount() As Long
    On Error GoTo errHandler
  CSLCount = colCSLDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.CSLCount"
End Function
Public Function colCSL() As Collection
    Set colCSL = colCSLDisplay
End Function



Public Function LoadPAY(EXCHID As String)
    On Error GoTo errHandler
Dim oPersist As c_Exchanges_P
    Set oPersist = New c_Exchanges_P
    Set colPayDisplay = Nothing
    Set colPayDisplay = New Collection
        SetSuperStatePAY oPersist.FetchPay(EXCHID)
    Set oPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.LoadPAY(EXCHID)", EXCHID
End Function
Friend Sub SetSuperStatePAY(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_Payment
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_Payment
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colPayDisplay.Add oDisp
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.SetSuperStatePAY(buffer)", buffer
End Sub
Public Function PAYItem(ByVal Index As Variant) As d_Payment
    On Error GoTo errHandler
    Set PAYItem = colPayDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.PAYItem(Index)", Index
End Function
Public Function PAYCount() As Long
    On Error GoTo errHandler
  PAYCount = colPayDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.PAYCount"
End Function
Public Function colPAY() As Collection
    Set colPAY = colPayDisplay
End Function
