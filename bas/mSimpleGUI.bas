Attribute VB_Name = "mSimpleGUI"
Public Sub SaveLayout(pG As TDBGrid, pFormName As String)
Dim i As Integer
    For i = 1 To pG.Columns.Count
        SaveSetting App.EXEName, pFormName, CStr(i), pG.Columns(i - 1).Width
    Next
End Sub
Public Sub SaveLayoutLvw(pG As ListView, pFormName As String)
Dim i As Integer
    For i = 1 To pG.ColumnHeaders.Count
        SaveSetting App.EXEName, pFormName, CStr(i), pG.ColumnHeaders(i).Width
    Next
End Sub

