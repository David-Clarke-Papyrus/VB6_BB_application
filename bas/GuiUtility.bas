Attribute VB_Name = "GuiUtility"

Option Explicit
Public frmWS As frmWaitStatus
Public Const DOCAPPROVAL = "You do not have authority to issue a document. Ask the supervisor to change your authorization level."
Public Const DOCACCESS = "You do not have authority to view a document. Ask the supervisor to change your authorization level."
Public gSTAFFID As Long
Public Const ISBNLENGTH = 13

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
    lpPoint As POINTAPI) As Long



Public Sub mSetfocus(ctl As Control)
On Error Resume Next
    ctl.Visible = True
    ctl.Enabled = True
    ctl.SetFocus
End Sub
Public Function OptionLoop(pCurrent As Long, pMax As Integer) As Long
Dim iNext As Integer
    iNext = pCurrent + 1
    If iNext > pMax Then iNext = 0
    OptionLoop = iNext
End Function
Public Function TranslateSince(pIn As Integer) As String
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
    Case 0
        TranslateSince = "&Since: Today"
    End Select
End Function
Public Sub LoadListbox(lb As ListBox, List As z_TextList)
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub

Public Sub LoadListboxSimple(lb As ListBox, List As z_TextListSimple)
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub

Public Sub LoadListboxColl(lb As ListBox, List As Collection)
Dim vntItem As Variant

    With lb
        .Clear
        For Each vntItem In List
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub

Public Sub LoadCombo(Combo As ComboBox, List As z_TextList, Optional iColumn As Integer)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            If iColumn > 0 Then
                .AddItem vntItem(iColumn)
            Else
                .AddItem vntItem(0)
            End If
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub
Public Sub LoadComboEx(Combo As EXCOMBOBOXLib.Items, List As z_TextList)
Dim vntItem As Variant
Dim itms As Items
    Set itms = Combo
    With itms
        .RemoveAllItems
        For Each vntItem In List
            itms.AddItem vntItem
        Next
       ' If .Count > 0 Then . = 0
    End With
    
End Sub

Public Sub AutoSelect(ctl As Control)
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
End Sub

Public Sub LoadComboFromTextListCol(Combo As Control, List As z_TextListCol)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub
Public Sub LoadComboFromTextListSimple(Combo As Control, List As z_TextListSimple)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
         '   MsgBox vntItem(0)
            .AddItem vntItem(0)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub

Public Sub LoadCombocol(Combo As ComboBox, List As Collection)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub

Public Function ParseMultiSelect(pCtrl As ListBox) As Collection
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
End Function

Public Function ISOpenForm(pType As Form, j As Integer) As Boolean
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
End Function


Public Sub SaveLayout(pG As TDBGrid, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
Dim i As Integer
    If Not pG Is Nothing Then
        For i = 1 To pG.Columns.Count
            SaveSetting "PBKS", pFormName, CStr(i), pG.Columns(i - 1).Width
        Next
    End If
    If Not IsMissing(pHeight) Then
        If pHeight > 0 Then
            SaveSetting "PBKS", pFormName, "Height", CStr(pHeight)
        End If
    End If
    If Not IsMissing(pWidth) Then
        If pWidth > 0 Then
            SaveSetting "PBKS", pFormName, "Width", CStr(pWidth)
        End If
    End If
            
End Sub
Public Sub SaveLayoutLvw(pG As ListView, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
Dim i As Integer
    For i = 1 To pG.ColumnHeaders.Count
        SaveSetting "PBKS", pFormName, CStr(i), pG.ColumnHeaders(i).Width
    Next
    If Not IsMissing(pHeight) Then
        If pHeight > 0 Then
            SaveSetting "PBKS", pFormName, "Height", CStr(pHeight)
        End If
    End If
    If Not IsMissing(pWidth) Then
        If pWidth > 0 Then
            SaveSetting "PBKS", pFormName, "Width", CStr(pWidth)
        End If
    End If
End Sub

Public Sub SetGridLayout(pG As TDBGrid, pFormName As String)
Dim i As Integer
    For i = 1 To pG.Columns.Count
        pG.Columns(i - 1).Width = GetSetting("PBKS", pFormName, CStr(i), pG.Columns(i - 1).Width)
    Next
End Sub
Public Sub SetLvwLayout(pG As ListView, pFormName As String)
Dim i As Integer
    For i = 1 To pG.ColumnHeaders.Count
        pG.ColumnHeaders(i).Width = GetSetting("PBKS", pFormName, CStr(i), pG.ColumnHeaders(i).Width)
    Next
End Sub

Public Function SetFormSize(f As Form)
Dim H As Long
Dim w As Long
    H = CLng(GetSetting("PBKS", f.Name, "Height", 0))
    w = CLng(GetSetting("PBKS", f.Name, "Width", 0))
    If H > 0 Then
        f.Height = H
    End If
    If w > 0 Then
        f.Width = w
    End If
End Function
Public Sub UnsetMenu()
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuAdjust.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuCreateCreditNote.Enabled = False
    'Forms(0).mnuProductPreview.Visible = False
    Forms(0).mnuSaveColumnWidths.Enabled = False
    Forms(0).mnuEmail.Enabled = False
    Forms(0).mnuEDI.Enabled = False
    Forms(0).mnuOutlook.Enabled = False
    Forms(0).mnuCopyLines.Enabled = False
    Forms(0).mnuPastelines.Enabled = False
    Forms(0).mnuDelact.Enabled = False
    Forms(0).mnuHeader.Enabled = False

End Sub

Public Sub WaitMsg(pMsg As String, pOn As Boolean, Optional frm As Form)
    If pOn Then
        If Not frm Is Nothing Then frm.Refresh
        Screen.MousePointer = vbHourglass
        Set frmWS = New frmWaitStatus
        frmWS.component pMsg
        frmWS.Show
        frmWS.Refresh
    Else
        If Not frmWS Is Nothing Then
            Unload frmWS
            Set frmWS = Nothing
        End If
        Screen.MousePointer = vbDefault
        If Not frm Is Nothing Then frm.Refresh
    End If
End Sub
'Public Sub HandleError()
'    On Error GoTo errHandler
'Dim strMsg As String
'Dim frmErr As frmError
'Dim strPos As String
'
'    If InException Then
'        MsgBox ErrDescription, vbOKOnly, "Exception"
'    Else
'        If ErrInIDE Then
'            frmShowError.ErrorReport = ErrReport
'        Else
'            Screen.MousePointer = vbDefault
'            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
'                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
'            Else
'                Select Case ErrNumber
'                    Case EXC_GENERAL:    strMsg = ErrDescription
'                    Case EXC_CANCELLED:  'nothing to do - it is silent exception.
'                    Case EXC_MULTIPLE:   strMsg = ErrDescription
'                    Case EXC_VALIDATION: strMsg = ErrDescription
'                End Select
'                Set frmErr = New frmError
'                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
'                frmErr.Show vbModal
'            End If
'            If frmWS Is Nothing Then
'            Else
'                Unload frmWS
'            End If
'        End If
'      '  MsgBox "Before ErrSaveToFile"
'        ErrSaveToFile
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "GUIUtility.HandleError"
'End Sub



Public Sub HandleErrorQuiet(pCLose As Boolean)
    pCLose = False
    On Error GoTo errHandler
    If InException Then
        MsgBox Err.Description, vbOKOnly, "Exception"
    Else
        ErrSaveToFile
        pCLose = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.HandleError"
End Sub

Public Sub LoadComboboxColl(Combo As ComboBox, List As Collection)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub

Public Function HandleTextWithBites(pText As String) As String
Dim strArg As String
Dim iStart As Integer
Dim iEnd As Integer
Dim oU As New z_UTIL
Dim strResult As String
Dim f As frmFindTextBite

    iStart = 0
    iEnd = 0
    iStart = InStr(1, pText, "?") + 1
    If iStart = 0 Then
        HandleTextWithBites = pText
        Exit Function
    End If
    HandleTextWithBites = pText
    strResult = ""
    iEnd = InStr(iStart, pText, "?")
    If iStart > 0 And iEnd > iStart Then
        strArg = Trim(Mid(pText, iStart, iEnd - iStart))
        strResult = oU.GetTextBite(strArg)
        If strResult > "" Then
            HandleTextWithBites = Replace(pText, "?" & strArg & "?", strResult)
        End If
    Else
    End If

End Function
Public Function mIsAmongBookmarks(X As XArrayDB, ID As Variant, G As TDBGrid, IDPosInGrid As Long, comptype As String) As Boolean
    Dim i As Integer
    mIsAmongBookmarks = False
    For i = 1 To G.SelBookmarks.Count
        If comptype = "LONG" Then
            If Val(X.Value(G.SelBookmarks(i - 1), IDPosInGrid)) = ID Then
                mIsAmongBookmarks = True
                Exit For
            End If
        Else
            If comptype = "UNIQUEIDENTIFIER" Or comptype = "STRING" Then
                If (X.Value(G.SelBookmarks(i - 1), IDPosInGrid)) = ID Then
                    mIsAmongBookmarks = True
                    Exit For
                End If
            End If
        End If
    Next i
End Function


' Get mouse X coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen

Function MouseX(Optional ByVal hwnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseX = lpPoint.X
End Function

' Get mouse Y coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen

Function MouseY(Optional ByVal hwnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseY = lpPoint.Y
End Function


