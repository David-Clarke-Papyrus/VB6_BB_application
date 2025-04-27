VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_DocumentControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolItems As Collection
Private mcolDeleted As Collection
Dim MaxKeyUsed As Long
Private mlngEditing As Long

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_DocumentControl
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems.Item((Index))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.Item(Index)", Index
End Function
Public Property Get FindDC(pDocCOde As String) As a_DocumentControl
    On Error GoTo errHandler
Dim oDC As a_DocumentControl
    For Each oDC In mcolItems
        If oDC.DOCCode = pDocCOde And oDC.WorkstationID = oPC.WorkstationID Then
            Set FindDC = oDC
            Exit For
        End If
    Next
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.FindDC(pDocCode)", pDocCOde
End Property
Public Function Add() As a_DocumentControl
    On Error GoTo errHandler
  Dim oDC As a_DocumentControl
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oDC = New a_DocumentControl
  With oDC
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
    MaxKeyUsed = MaxKeyUsed + 1
    oDC.Key = MaxKeyUsed & "k"
  End With
  Set Add = oDC
  Set oDC = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.Add"
End Function

Friend Sub AddDC(pDC As a_DocumentControl)
    On Error GoTo errHandler

  mcolItems.Add pDC, pDC.Key

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.AddDC(pDC)", pDC
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  mcolDeleted.Add mcolItems(Index)
  mcolItems.Remove Index
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.Remove(Index)", Index
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oDC As a_DocumentControl
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oDC In mcolItems
      If oDC.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.IsDirty"
End Function

Public Sub BeginEdit()
    On Error GoTo errHandler
Dim oDC As a_DocumentControl
  
  For Each oDC In mcolItems
    oDC.BeginEdit
  Next
  For Each oDC In mcolDeleted
    oDC.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
Dim oDC As a_DocumentControl
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oDC = mcolItems(lngIndex)
        If oDC.IsNew Then
            oDC.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oDC.CancelEdit
        End If
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oDC = mcolDeleted(lngIndex)
      With oDC
          .CancelEdit
        If Not .IsNew Then
          If Not .IsDeleted Then
            mcolItems.Add oDC, oDC.Key
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
        If oDC.IsEditing Then oDC.CancelEdit
      Set oDC = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oDC As a_DocumentControl

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oDC In mcolItems
    oDC.ApplyEdit
  Next
  For Each oDC In mcolDeleted
    oDC.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.ApplyEdit"
End Sub
Public Sub Delete()
    On Error GoTo errHandler
  Dim lngIndex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Remove lngIndex
  Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oDC As a_DocumentControl
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oDC In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oDC.GetSuperState
    Next
    For Each oDC In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oDC.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDC As a_DocumentControl
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oDC = New a_DocumentControl
      oDC.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oDC.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oDC.BeginEdit
      Next
      If Not oDC.IsDeleted Then
        mcolItems.Add oDC, oDC.Key
      Else
        mcolDeleted.Add oDC
      End If
      Set oDC = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.SetSuperState(buffer)", buffer
End Sub

Public Sub VerifyDocumentTypes(pWorkstation As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim WSID As Long
Dim OpenResult As Integer

    WSID = oPC.Configuration.Workstations.Key(pWorkstation)
    Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_VerifyDocumentTypes"
    cmd.CommandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , WSID)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@PrinterID", adInteger, adParamInput, , oPC.Configuration.DefaultPrinterID)
    cmd.Parameters.Append prm
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DocumentControl.VerifyDocumentTypes(pWorkstation)", pWorkstation
End Sub

