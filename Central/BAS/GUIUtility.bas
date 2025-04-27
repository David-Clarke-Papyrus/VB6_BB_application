Attribute VB_Name = "GUIUTility"
Option Explicit

Public frmWS As frmWaitStatus
Public Const ISBNLENGTH = 13

Public Const DOCAPPROVAL = "You do not have authority to issue a document. Ask the supervisor to change your authorization level."
Public gSTAFFID As Long
Public Sub mSetfocus(ctl As Control)
    ctl.Visible = True
    ctl.Enabled = True
    ctl.SetFocus
End Sub

Public Function OptionLoop(pCurrent As Long, pMax As Integer) As Long
Dim iNext As Integer
    iNext = pCurrent + 1
    If iNext > pMax Then iNext = 1
    OptionLoop = iNext
End Function
Public Function OptionLoopStores(pCurrent As Long, pMax As Integer, pExcludeLocal As Boolean) As Long
Dim iNext As Integer
    iNext = pCurrent + 1
    If iNext > pMax Then iNext = 1
    If pExcludeLocal Then
        If oPC.Configuration.Stores(iNext) Is oPC.Configuration.DefaultStore Then
            If iNext = 1 And pMax = 1 Then
                iNext = 0
            Else
                iNext = iNext + 1
            End If
        End If
    End If
    OptionLoopStores = iNext
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

Public Sub LoadListboxSimple(lb As ListBox, List As z_TextListSImple)
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


Public Sub LoadComboEx(Combo As EXCOMBOBOXLib.Items, List As z_TextList)
Dim vntItem As Variant
Dim itms As Items
    Set itms = Combo
    With itms
        .RemoveAllItems
        For Each vntItem In List
            itms.AddItem vntItem
        Next
    End With
    
End Sub

Public Sub AutoSelect(ctl As Control)
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
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
            SaveSetting "CENTRAL", pFormName, CStr(i), pG.Columns(i - 1).Width
        Next
    End If
    If Not IsMissing(pHeight) Then
        If pHeight > 0 Then
            SaveSetting "CENTRAL", pFormName, "Height", CStr(pHeight)
        End If
    End If
    If Not IsMissing(pWidth) Then
        If pWidth > 0 Then
            SaveSetting "CENTRAL", pFormName, "Width", CStr(pWidth)
        End If
    End If
            
End Sub

Public Sub WaitMsg(pMsg As String, pON As Boolean, Optional frm As Form)
    If pON Then
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

Public Sub SetGridLayout(pG As TDBGrid, pFormName As String)
Dim i As Integer
    For i = 1 To pG.Columns.Count
        pG.Columns(i - 1).Width = GetSetting("CENTRAL", pFormName, CStr(i), pG.Columns(i - 1).Width)
    Next
End Sub

Public Function SetFormSize(f As Form)
Dim H As Long
Dim w As Long
    H = CLng(GetSetting("CENTRAL", f.Name, "Height", 0))
    w = CLng(GetSetting("CENTRAL", f.Name, "Width", 0))
    If H > 0 Then
        f.Height = H
    End If
    If w > 0 Then
        f.Width = w
    End If
End Function

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
Public Sub UnsetMenu()
On Error Resume Next
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

