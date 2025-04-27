Attribute VB_Name = "mFormLayout"
Option Explicit
'Public Sub SaveSplits(pFormName As String, CM As ControlManager)
'    On Error GoTo errHandler
'Dim i As Long
'Dim rw As String
'Dim Max As Integer
'    Max = CM.Splitters.Count
'    For i = 0 To Max - 1
'        rw = CStr(IIf(CM.Splitters(i).Orientation = orVertical, CM.Splitters(i).Xc, CM.Splitters(i).Yc))
'        SaveSetting "PBKS", pFormName & CStr(i), "Dims", rw
'    Next
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "mFormLayout.SaveSplits(pFormName,CM)", Array(pFormName, CM)
'End Sub
'Public Sub SetCM(f As Form, CM As ControlManager)
'    On Error GoTo errHandler
'Dim i As Long
'Dim rw As String
'Dim arDims() As String
'    For i = 0 To CM.Splitters.Count - 1
'        CM.MoveSplitter i, CLng(GetSetting("PBKS", f.Name & CStr(i), "Dims", "1000"))
'    Next
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "mFormLayout.SetCM(f,CM)", Array(f, CM)
'End Sub

Public Sub SaveFormSize(pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
    On Error GoTo errHandler
Dim i As Integer
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
            
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SaveFormSize(pFormName,pHeight,pWidth)", Array(pFormName, pHeight, pWidth)
End Sub

Public Sub SaveLayout(pG As TDBGrid, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
    On Error GoTo errHandler
Dim i As Integer
    If Not pG Is Nothing Then
        For i = 1 To pG.Columns.Count
            SaveSetting "PBKS", pFormName, CStr(i), CStr(pG.Columns(i - 1).Width)
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
            
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SaveLayout(pG,pFormName,pHeight,pWidth)", Array(pG, pFormName, pHeight, _
         pWidth)
End Sub
Public Sub SaveLayoutLvw(pG As ListView, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SaveLayoutLvw(pG,pFormName,pHeight,pWidth)", Array(pG, pFormName, pHeight, _
         pWidth)
End Sub

Public Sub SetGridLayout(pG As TDBGrid, pFormName As String)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To pG.Columns.Count
        pG.Columns(i - 1).Width = CLng(GetSetting("PBKS", pFormName, CStr(i), pG.Columns(i - 1).Width))
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SetGridLayout(pG,pFormName)", Array(pG, pFormName)
End Sub
Public Sub SetLvwLayout(pG As ListView, pFormName As String)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To pG.ColumnHeaders.Count
        pG.ColumnHeaders(i).Width = GetSetting("PBKS", pFormName, CStr(i), pG.ColumnHeaders(i).Width)
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SetLvwLayout(pG,pFormName)", Array(pG, pFormName)
End Sub

Public Function SetFormSize(f As Form)
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SetFormSize(f)", f
End Function
Public Sub UnsetMenu()
    On Error GoTo errHandler
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.UnsetMenu"
End Sub


Public Sub SaveContourCubeLayout(ltFile As String, pCC As CCubeX2.ContourCubeX)
    On Error GoTo errHandler
'Saving layout procedure
  Dim rsFields, Axis, Object, bInvertFilterSelection, Value, i, j, viewTotalsState, _
      viewGTotalsState, strExpand, fs
  rsFields = Array("Object", "Name", "Property", "Value")
  'Create an ADO recordset with 4 fields:
  Dim rscube As New ADODB.Recordset
  rscube.Fields.Append rsFields(0), adBSTR, 10
  rscube.Fields.Append rsFields(1), adBSTR, 50
  rscube.Fields.Append rsFields(2), adVariant, 50
  rscube.Fields.Append rsFields(3), adVariant, 255
  rscube.Open
  rscube.AddNew rsFields, Array("Cube", pCC.Name, "RootAxis", pCC.Cube.RootAxis)
  With pCC
    'Populate recordset with layout properties
    For Each Object In .Facts
      'Fact visibility
      rscube.AddNew rsFields, Array("Fact", Object.Name, "Visible", Object.Visible)
    Next
    For i = 0 To 1
        If i = 0 Then Set Axis = .VAxis Else Set Axis = .HAxis
        For Each Object In Axis.Dims
          'Dimension positions and properties
          rscube.AddNew rsFields, Array("Dim", Object.Name, "Axis", Object.CubeDim.Axis)
          rscube.AddNew rsFields, Array("Dim", Object.Name, "Pos", Object.CubeDim.pos)
        Next
    Next
    For Each Object In .Dims
        rscube.AddNew rsFields, Array("Dim", Object.Name, "Totals", Object.NoTotals)
        rscube.AddNew rsFields, Array("Dim", Object.Name, "Descending", Object.Descending)
        'Dimension filters:
        'To minimize the file, choose the minimum set between hidden and visible
        'values to save
        bInvertFilterSelection = (Object.CubeDim.GetValues(2).Count > Object.CubeDim.GetValues(1).Count)
        rscube.AddNew rsFields, Array("DimsFilter", "InvertFilterSelection", Object.Name, bInvertFilterSelection)
        For Each Value In Object.CubeDim.GetValues(IIf(bInvertFilterSelection, 1, 2))
          rscube.AddNew rsFields, Array("DimsFilter", "Filter", Object.Name, Value)
        Next
    Next
    'Save axis expand states
    'Temporarily turn off totals, in order not to save sections that
    'correspond to dimension totals
    viewTotalsState = .NoTotals
    viewGTotalsState = .NoGrandTotals
    .NoTotals = True
    .NoGrandTotals = True
    'Cycle through all sections on both axes and save their state
    If .HAxis.Length > 0 Then
      For i = 0 To .HAxis.Length - 1
        strExpand = ""
        For j = 0 To .HAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .HAxis.GetSection(i).getValue(j)
        Next j
        rscube.AddNew rsFields, Array("Axis", "Horizontal", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    If .VAxis.Length > 0 Then
      For i = 0 To .VAxis.Length - 1
        strExpand = ""
        For j = 0 To .VAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .VAxis.GetSection(i).getValue(j)
        Next j
        rscube.AddNew rsFields, Array("Axis", "Vertical", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    'Restore view totals
    .NoTotals = viewTotalsState
    .NoGrandTotals = viewGTotalsState
  End With
  'Verify if the file already exists and eventually delete it before saving
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.FileExists(ltFile) Then fs.DeleteFile (ltFile)
  rscube.Save ltFile, adPersistXML
  rscube.Close
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.SaveContourCubeLayout(ltFile,pCC)", Array(ltFile, pCC)
End Sub

Sub LoadContourcubeLayout(ltFile As String, pCC As CCubeX2.ContourCubeX)
    On Error GoTo errHandler
'Loading layout procedure
  Dim FactSettings, DimSettings, Object, DimFilters, AxisSettings, i, bInvertFilterSelection
  Dim rscube As New ADODB.Recordset
  'First open the saved XML layout file
  rscube.Open ltFile
  With pCC
    'Restore cube properties
    rscube.MoveFirst
    rscube.Filter = "Object='Cube'"
    .Cube.RootAxis = CInt(rscube.GetRows()(3, 0))
    'Fact visibility
    rscube.Filter = "Object='Fact'"
    FactSettings = rscube.GetRows()
    For i = 0 To UBound(FactSettings, 2)
      If LCase(CStr(FactSettings(2, i))) = "visible" Then
        If .Facts.Exists(CStr(FactSettings(1, i))) Then _
           .Facts(CStr(FactSettings(1, i))).Visible = CBool(FactSettings(3, i))
      End If
    Next i
    'Set up dimension positions, totalling and sort orders
    rscube.Filter = "Object='Dim'"
    DimSettings = rscube.GetRows()
    For Each Object In .Dims
        If Object.CubeDim.Axis <> xda_invisible Then Object.CubeDim.MoveTo xda_outside
    Next
    For i = 0 To UBound(DimSettings, 2)
      If .Dims.Exists(CStr(DimSettings(1, i))) Then
        Select Case LCase(CStr(DimSettings(2, i)))
        Case "axis":
          .Dims(CStr(DimSettings(1, i))).CubeDim.MoveTo CInt(DimSettings(3, i))
        Case "pos":
          .Dims(CStr(DimSettings(1, i))).CubeDim.MoveTo .Dims(CStr(DimSettings(1, i))).CubeDim.Axis, CInt(DimSettings(3, i))
        Case "totals":
          .Dims(CStr(DimSettings(1, i))).NoTotals = CBool(DimSettings(3, i))
        Case "descending":
          .Dims(CStr(DimSettings(1, i))).Descending = CBool(DimSettings(3, i))
        End Select
      End If
    Next i
    If .Dims.Count = 0 Then Exit Sub
    .Active = True
    'Dimension filter states
    rscube.Filter = "Object='DimsFilter'"
    DimFilters = rscube.GetRows()
    For i = 0 To UBound(DimFilters, 2)
      If .Dims.Exists(CStr(DimFilters(2, i))) Then
        Select Case LCase(CStr(DimFilters(1, i)))
        Case "invertfilterselection":
          bInvertFilterSelection = CBool(DimFilters(3, i))
          .Dims(CStr(DimFilters(2, i))).CubeDim.Filter IIf(bInvertFilterSelection, xfo_FilterAll, xfo_Reset)
        Case "filter":
          .Dims(CStr(DimFilters(2, i))).CubeDim.FilterValue DimFilters(3, i), Not bInvertFilterSelection
        End Select
      End If
    Next i
    .Cube.DimensionsFilter.Apply
    'Finally, restore expand status of each axis section
    .HAxis.DrillDownLevel = .HAxis.Width - 1
    .VAxis.DrillDownLevel = .VAxis.Width - 1
    rscube.Filter = "Object='Axis'"
    AxisSettings = rscube.GetRows()
    For i = 0 To UBound(AxisSettings, 2)
      ExpandSection CStr(AxisSettings(1, i)), CStr(AxisSettings(3, i)), pCC
    Next i
  End With
  rscube.Close
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.LoadContourcubeLayout(ltFile,pCC)", Array(ltFile, pCC)
End Sub

Sub ExpandSection(strAxis As String, strExpand As String, pCC As CCubeX2.ContourCubeX)
    On Error GoTo errHandler
    
    On Error Resume Next
'This procedure restores saved state of an axis section
'It searches for given combination of dim values along the axis,
'and expands the section found
  Dim Axis As IViewAxis, i, j, aExpand
  aExpand = Split(strExpand, Chr(10))
  If LCase(strAxis) = "horizontal" Then Set Axis = pCC.HAxis Else Set Axis = pCC.VAxis
  i = 0
  Do While i < Axis.Length
    j = 0
    Do While j <= UBound(aExpand, 1)
      If CStr(Axis.GetSection(i).getValue(j)) <> aExpand(j) Then Exit Do
      j = j + 1
    Loop
    If j > UBound(aExpand, 1) Then Exit Do
    i = i + 1
  Loop
  If i < Axis.Length Then Axis.GetSection(i).Collapse UBound(aExpand, 1), True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mFormLayout.ExpandSection(strAxis,strExpand,pCC)", Array(strAxis, strExpand, pCC)
End Sub


