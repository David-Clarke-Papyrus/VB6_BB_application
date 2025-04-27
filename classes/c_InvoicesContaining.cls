VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_InvoicesContaining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colDisplay As Collection


Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set colDisplay = New Collection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = colDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = colDisplay.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As d_InvoicesContaining
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = colDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining.Item(Index)", Index
End Function

Public Sub Load(pTPID As Long, pCODE As String)
    On Error GoTo errHandler
Dim oPersist As c_InvoicesContaining_P

    Set oPersist = New c_InvoicesContaining_P
    SetSuperState oPersist.Fetch(pTPID, pCODE)
    Set oPersist = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining.Load(pTPID)", pTPID
End Sub

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim odNS As d_InvoicesContaining
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odNS = New d_InvoicesContaining
        odNS.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add odNS
        Set odNS = Nothing
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining.SetSuperState(buffer)", buffer
End Sub



