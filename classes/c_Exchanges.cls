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
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As ExchangeProps
Private colDisplay As Collection

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

Public Function Load(pOPID As String, pZID As String, pEID As String)
    On Error GoTo errHandler
Dim oPersist As c_Exchanges_P
    Set oPersist = New c_Exchanges_P
    If pOPID > "" Then
        SetSuperState oPersist.Fetch(pOPID, "", "")
    ElseIf pZID > "" Then
        SetSuperState oPersist.Fetch("", pZID, "")
    Else
        SetSuperState oPersist.Fetch("", "", pEID)
    End If
    Set oPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges.Load(pOPID,pZID)", Array(pOPID, pZID)
End Function
Public Function LoadCustomerRequests(pOSorALL As String, dSince As Date)
    On Error GoTo errHandler
Dim oPersist As c_Exchanges_P
    Set oPersist = New c_Exchanges_P
    SetSuperState oPersist.FetchCustomequests(pOSorALL, dSince)
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



