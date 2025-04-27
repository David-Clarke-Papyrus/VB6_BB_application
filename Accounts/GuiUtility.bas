Attribute VB_Name = "GuiUtility"

Option Explicit
Public Const DOCAPPROVAL = "You do not have authority to issue a document. Ask the supervisor to change your authorization level."
Public Const DOCACCESS = "You do not have authority to view a document. Ask the supervisor to change your authorization level."
Public gSTAFFID As Long
Public Const ISBNLENGTH = 13

Private Type POINTAPI
    x As Long
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
Public Function TranslateSince(pIN As Integer) As String
    Select Case pIN
    Case 2
        TranslateSince = "Last week"
    Case 3
        TranslateSince = "Last month"
    Case 4
        TranslateSince = "Last quarter"
    Case 5
        TranslateSince = "Last year"
    Case 1
        TranslateSince = "<any date>"
    Case 0
        TranslateSince = "Today"
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
'Public Sub LoadComboEx(Combo As Object, List As z_TextList)
'Dim vntItem As Variant
'Dim itms As Items
''EXCOMBOBOXLib.Items
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
Public Sub LoadListboxFromRecordset(lb As ListBox, rs As ADODB.Recordset)
Dim vntItem As Variant

    With lb
        .Clear
        Do While Not rs.EOF
            .AddItem rs.Fields(0), CStr(rs.Fields(1))
        Loop
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub
'Public Sub LoadComboFromRecordset(Combo As ComboBox, rs As ADODB.Recordset)
'Dim vntItem As Variant
'Dim itms As Items
'
'    With Combo
'        .Clear
'
'        Do While Not rs.EOF
'            .AddItem rs.Fields(0)
'            rs.MoveNext
'
'        Loop
'        If .ListCount > 0 Then .ListIndex = 0
'    End With
'
'End Sub

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

'Public Function HandleTextWithBites(pText As String) As String
'Dim strArg As String
'Dim iStart As Integer
'Dim iEnd As Integer
'Dim oU As New z_UTIL
'Dim strResult As String
'Dim f As frmFindTextBite
'
'    iStart = 0
'    iEnd = 0
'    iStart = InStr(1, pText, "?") + 1
'    If iStart = 0 Then
'        HandleTextWithBites = pText
'        Exit Function
'    End If
'    HandleTextWithBites = pText
'    strResult = ""
'    iEnd = InStr(iStart, pText, "?")
'    If iStart > 0 And iEnd > iStart Then
'        strArg = Trim(Mid(pText, iStart, iEnd - iStart))
'        strResult = oU.GetTextBite(strArg)
'        If strResult > "" Then
'            HandleTextWithBites = Replace(pText, "?" & strArg & "?", strResult)
'        End If
'    Else
'    End If
'
'End Function
Public Function mIsAmongBookmarks(x As XArrayDB, ID As Variant, G As TDBGrid, IDPosInGrid As Long, comptype As String) As Boolean
    Dim i As Integer
    mIsAmongBookmarks = False
    For i = 1 To G.SelBookmarks.Count
        If comptype = "LONG" Then
            If val(x.Value(G.SelBookmarks(i - 1), IDPosInGrid)) = ID Then
                mIsAmongBookmarks = True
                Exit For
            End If
        Else
            If comptype = "UNIQUEIDENTIFIER" Or comptype = "STRING" Then
                If (x.Value(G.SelBookmarks(i - 1), IDPosInGrid)) = ID Then
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
    MouseX = lpPoint.x
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


