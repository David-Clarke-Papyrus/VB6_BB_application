Attribute VB_Name = "GuiUtility"

Option Explicit
Public frmWS As frmWaitStatus

Public Const DOCAPPROVAL = "You do not have authority to issue a document. Ask the supervisor to change your authorization level."

Public Function OptionLoop(pCurrent As Long, pMax As Integer) As Long
    On Error GoTo errHandler
Dim iNext As Integer
    iNext = pCurrent + 1
    If iNext > pMax Then iNext = 1
    OptionLoop = iNext
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.OptionLoop(pCurrent,pMax)", Array(pCurrent, pMax)
End Function
Public Function TranslateSince(pIn As Integer) As String
    On Error GoTo errHandler
    Select Case pIn
    Case 2
        TranslateSince = "&Since: Last week"
    Case 3
        TranslateSince = "&Since: Last month"
    Case 4
        TranslateSince = "&Since: Last quarter"
    Case 5
        TranslateSince = "&Since: Last year"
    Case 1
        TranslateSince = "&Since: <any date>"
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.TranslateSince(pIn)", pIn
End Function
Public Sub LoadListboxSimple(lb As ListBox, List As z_TextListSimple)
    On Error GoTo errHandler
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.LoadListboxSimple(lb,List)", Array(lb, List)
End Sub
Public Sub LoadComboSimple(lb As ComboBox, List As z_TextListSimple)
    On Error GoTo errHandler
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.LoadComboSimple(lb,List)", Array(lb, List)
End Sub

Public Sub LoadListboxColl(lb As ListBox, List As Collection)
    On Error GoTo errHandler
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.LoadListboxColl(lb,List)", Array(lb, List)
End Sub

'Public Sub LoadCombo(Combo As ComboBox, List As z_TextList, Optional iColumn As Integer)
'Dim vntItem As Variant
'
'    With Combo
'        .Clear
'        For Each vntItem In List
'            If iColumn > 0 Then
'                .AddItem vntItem(iColumn)
'            Else
'                .AddItem vntItem(0)
'            End If
'        Next
'        If .ListCount > 0 Then .ListIndex = 0
'    End With
'
'End Sub
'Public Sub LoadComboEx(Combo As EXCOMBOBOXLib.Items, List As z_TextList)
'Dim vntItem As Variant
'Dim itms As Items
'    Set itms = Combo
'    With itms
'        .RemoveAllItems
'        For Each vntItem In List
'            itms.AddItem vntItem
'        Next
'       ' If .Count > 0 Then . = 0
'    End With
'
'End Sub

Public Sub AutoSelect(ctl As Control)
    On Error GoTo errHandler
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.AutoSelect(ctl)", ctl
End Sub

'Public Sub LoadComboFromTextListCol(Combo As Control, List As z_TextListCol)
'Dim vntItem As Variant
'
'    With Combo
'        .Clear
'        For Each vntItem In List
'            .AddItem vntItem
'        Next
'        If .ListCount > 0 Then .ListIndex = 0
'    End With
'
'End Sub

Public Sub LoadCombocol(Combo As ComboBox, List As Collection)
    On Error GoTo errHandler
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.LoadCombocol(Combo,List)", Array(Combo, List)
End Sub

Public Function ParseMultiSelect(pCtrl As ListBox) As Collection
    On Error GoTo errHandler
Dim Col As Collection
Dim i As Integer
    Set Col = New Collection
   For i = 0 To pCtrl.ListCount - 1
      If pCtrl.Selected(i) Then
         Col.Add pCtrl.List(i)
      End If
   Next i
    Set ParseMultiSelect = Col
    Set Col = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.ParseMultiSelect(pCtrl)", pCtrl
End Function

Public Function ISOpenForm(pType As Form, j As Integer) As Boolean
    On Error GoTo errHandler
Dim bFound As Boolean
Dim i As Integer
    bFound = False
    For i = 0 To Forms.Count
        If Forms(i).Name = pType.Name Then
            j = i
            bFound = True
            Exit For
        End If
    Next i
    ISOpenForm = bFound
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.ISOpenForm(pType,j)", Array(pType, j)
End Function


'Public Sub SaveLayout(pG As TDBGrid, pFormName As String)
'Dim i As Integer
'    For i = 1 To pG.Columns.Count
'        SaveSetting "PBKS", pFormName, CStr(i), pG.Columns(i - 1).Width
'    Next
'End Sub

Public Sub UnsetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuGenDisc.Enabled = False
    Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuAdjust.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSetPrinter.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.UnsetMenu"
End Sub

'Public Function SecurityControl(pLevel As Integer, pSTAFFID As Long, Optional pCancelled As Boolean, Optional pMsg As String, Optional pErrMsg As String) As Boolean
'Dim frmS As New frmSecurity
'    frmS.Component pMsg
'    frmS.Show vbModal
'    If frmS.Cancelled Then
'       ' pCancelled = True
'        SecurityControl = False
'        Exit Function
'    End If
'    Unload frmS
'    SecurityControl = True
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode, , pSTAFFID) < pLevel Then
'        If pErrMsg = "" Then pErrMsg = "You do not have security authority."
'        MsgBox pErrMsg, vbExclamation, "Action denied"
'        SecurityControl = False
'    End If
'End Function
Public Sub WaitMsg(pMsg As String, pON As Boolean, frm As Form)
    On Error GoTo errHandler
    If pON Then
        frm.Refresh
        Screen.MousePointer = vbHourglass
        Set frmWS = New frmWaitStatus
        frmWS.Component pMsg
        frmWS.Show
        frmWS.Refresh
    Else
        If Not frmWS Is Nothing Then
            Unload frmWS
            Set frmWS = Nothing
        End If
        Screen.MousePointer = vbDefault
        frm.Refresh
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.WaitMsg(pMsg,pON,frm)", Array(pMsg, pON, frm)
End Sub

Public Sub HandleError()
    On Error GoTo errHandler
    If InException Then
        MsgBox ErrDescription, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            MsgBox "An error has occurred. The text of the message is stored in " & App.Path & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport, vbInformation, "Error in application"
            If frmWS Is Nothing Then
            Else
                Unload frmWS
            End If
        End If
        ErrSaveToFile
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.HandleError"
End Sub
Public Sub HandleErrorQuiet(pCLose As Boolean)
    pCLose = False
    On Error GoTo errHandler
'    If InException Then
'        oTF.WriteToTextFile ErrDescription
'    '    MsgBox ErrDescription, vbOKOnly, "Exception"
'    Else
        ErrSaveToFile
        pCLose = True
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.HandleError"
End Sub
