VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CCA2C66D-33FD-11D5-8D72-005004532BDF}#1.3#0"; "CCubeX.ocx"
Begin VB.Form frmProductPT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Documents per supplier"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   11880
   Begin CCubeX.ContourCubeX CC 
      Height          =   5520
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   11535
      BackColor       =   14737632
      Enabled         =   -1  'True
      MainAxis        =   0
      DataSourceType  =   0
      ConnectionString=   ""
      SQL             =   ""
      PreGrouping     =   -1  'True
      Active          =   0   'False
      HDrillDownLevel =   -1
      VDrillDownLevel =   1
      Transposed      =   0   'False
      SuppressZeroRows=   0   'False
      SuppressZeroCols=   0   'False
      ViewFlags       =   0
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowFilter     =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      AllowTitle      =   0   'False
      ShowAsPercent   =   0
      TotalsString    =   ""
      CubeTitle       =   ""
      TitleAlign      =   0
      TitleBkColor    =   13160660
      DimBkColor      =   13160660
      DimTitleBkColor =   6956042
      DimTitleInactiveBkColor=   8421504
      DimFilterBkColor=   13160660
      InactiveDimAreaBkColor=   13160660
      HeadingBkColor  =   13160660
      DataGridColor   =   13160660
      DataBkColor     =   16777215
      TotalBkColor    =   14679807
      GrandTotalBkColor=   14679807
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimFilterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DataFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TotalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GrandTotalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
      Object.Visible         =   -1  'True
      MousePointer    =   0
      TitleForeColor  =   0
      DimForeColor    =   0
      DimTitleForeColor=   16777215
      DimFilterForeColor=   0
      HeadingForeColor=   0
      DataForeColor   =   0
      TotalForeColor  =   -2147483640
      GrandTotalForeColor=   -2147483640
      UnusedDataAreaColor=   -2147483643
      MainAxisDim     =   ""
      DimTitleDragBkColor=   32768
      FactsCaption    =   "Facts"
      ShowFactsBitmap =   -1  'True
      ADOCursorLocation=   2
      AutoRefreshView =   0   'False
      FPErrString     =   "FPErr"
      NULLValueString =   ""
      NonExistentValueString=   ""
      DefaultFactFormat=   "###,###,###,###,###,##0.00"
      AllowFactFilter =   -1  'True
      VERSION_NO      =   2
      FIELDS_SETTINGS =   $"frmProductPT.frx":0000
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   10380
      Top             =   5955
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8955
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5670
      Width           =   1260
   End
   Begin VB.TextBox txtCubeName 
      Height          =   285
      Left            =   1845
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5610
      Width           =   3315
   End
   Begin VB.CommandButton cmdPrintCube 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5685
      Width           =   1260
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5910
      Width           =   1260
   End
   Begin VB.CommandButton cmdSaveCube 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5910
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5670
      Width           =   1260
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5670
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "'C' sorts SHIFT+C reverses sort, 'X' cancels sorting"
      Height          =   270
      Left            =   5385
      TabIndex        =   8
      Top             =   6210
      Width           =   3675
   End
End
Attribute VB_Name = "frmProductPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
    Caption = "Sales patterns"

End Sub

Private Sub CC_KeyUp(ByVal KeyCode As Long, ByVal Shift As Long)
    If KeyCode = vbKeyC Then
            If Shift = 1 Then
                CC.ViewFlags = CC.ViewFlags + xfDescending
            Else
                If CC.ViewFlags = xfDescending Then
                    CC.ViewFlags = CC.ViewFlags - xfDescending
                End If
            End If
            CC.SortByFact xda_vertical
    End If
    If KeyCode = vbKeyX Then
            CC.CancelFactSorting xda_vertical
            
            CC.DimFlags("Acno") = 0
            CC.ViewFlags = 0
    End If
End Sub

'Private Sub CCo_KeyUp(ByVal keycode As Long, ByVal Register As Long)
'  Dim Col As Long
'  Dim Row As Long
'  With CC
''   Press "C" key for sorting current column
'    If keycode = vbKeyC And .VAxis.Dims.Count > 0 Then
'        With .CurrentCell
'            If Register = 1 Then
'                CC.Facts(.Col).Descending = Not CC.Dims(.Col).Descending
'            End If
'            CC.SortGridByFact xda_vertical, .Col, .Row
'        End With
''   Press "R" key for sorting current row
'    ElseIf keycode = vbKeyR And .HAxis.Dims.Count > 0 Then
'        With .CurrentCell
'            CC.SortGridByFact xda_horizontal, .Col, .Row
'        End With
''   Press "A" key for abandon sorting
'    ElseIf keycode = vbKeyA Then
'      .CancelFactSorting
'    End If
'  End With
'End Sub
'
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    CC.Active = False
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim SQL As String

    CC.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    WaitMsg "Loading the pivot table . . . ", True, Me
    CC.DataSourceType = xcdt_Recordset
    CC.Open rs
   
    CC.Active = True
    WaitMsg "", False, Me
    Me.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdConnect_Click()
    ConnectToData
End Sub

Private Sub cmdLoad_Click()
Dim fs As New FileSystemObject
    CD1.DefaultExt = ".txt"
    CD1.DialogTitle = "Load stored cube"
    CD1.InitDir = "C:\PBKS\BU"
    
    CD1.ShowOpen
    txtCubeName = CD1.FileName
    If fs.FileExists(txtCubeName) Then
        CC.LoadCube txtCubeName
    Else
        MsgBox "Nothing to load"
    End If

End Sub

Private Sub cmdPrint_Click()
Dim res As Boolean
Dim fs As New FileSystemObject

    CC.ExportToFile oPC.SharedFolderRoot & "\HTML\SalesPatterns.html", oPC.SharedFolderRoot & "\HTML\SalesPatterns.html", xet_html
    MsgBox "Exported to file " & oPC.SharedFolderRoot & "\HTML\SalesPatterns.html"
    Exit Sub

End Sub

Private Sub cmdPrintCube_Click()
    CC.AllowTitle = True
    CC.CubeTitle = "Sales patterns: printed " & Format(Now(), "DD/mm/yyyy HH:HH AM/PM")
  '  CC.CubeFooter = "TEST Footer"

    CC.PrintCube True
    
End Sub

Private Sub cmdSaveCube_Click()
Dim fs As New FileSystemObject

    CD1.DefaultExt = ".txt"
    CD1.DialogTitle = "Save cube"
    CD1.InitDir = "C:\PBKS\BU"
    
    CD1.ShowOpen
    txtCubeName = CD1.FileName
    If fs.FileExists(txtCubeName) Then
        fs.DeleteFile txtCubeName
    End If
    CC.SaveCube txtCubeName
End Sub

'Private Sub cmdPrint_Click()
'    On Error GoTo errHandler
'    CC.ExportToFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", xet_html
'    CC.PrintCube True, False
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSImple
top = 400
left = 20
Width = 11900
Height = 6800

    
'    CC.Cube.Dims.Clear
'    CC.Cube.BaseFacts.Clear
'    CC.Cube.Facts.Clear
'    CC.Cube.Dims.Add "yr", "Yr", 5, xda_vertical
'    CC.Cube.Dims.Add "mth", "Mth", 5, xda_vertical
'    CC.Cube.Dims.Add "wk", "Wk", 5, xda_outside
'    CC.Cube.Dims.Add "BIC", "BIC", xoft_String, xda_outside
'    CC.Cube.Dims.Add "Br", "Br", xoft_String, xda_vertical
'    CC.Cube.Dims.Add "Acno", "Acno", xoft_String, xda_outside
'    CC.Cube.Dims.Add "ProductType", "ProductType", 5, xda_outside
'    CC.Cube.Dims.Add "EXCHANGENUMBER", "EXCHANGENUMBER", xoft_String, xda_outside
'    CC.Cube.Dims.Add "EAN", "EAN", xoft_String, xda_outside
'    CC.Cube.BaseFacts.Add "bfQty", "QTY" ', xfaa_SUM + xfam_RANKA
'    CC.Cube.BaseFacts.Add "bfVal", "VAL" ' , xfaa_SUM
'    CC.Cube.Facts.Add "fQty", "bfQty", xfaa_SUM
'    CC.Facts("fQty").Appearance.Format = "### ### ##0.00"
'    CC.Facts("fQty").Visible = True
'    CC.Cube.Facts.Add "fVal", "bfVal", xfaa_SUM
'    CC.Facts("fVal").Appearance.Format = "### ### ##0.00"
'    CC.Facts("fVal").Visible = True
'    CC.Active = False
'    DoEvents
'    Screen.MousePointer = vbHourglass
'   ' CC.Cube.DataSourceType = xcdt_Recordset
'    rs.MoveFirst
'    If Not rs.EOF Then
'        CC.Cube.Open rs, True
'        'CC.SortByFact xda_vertical, 1
'       ' CC.Active = True
'    Else
'        MsgBox "No records", , "Status"
'    End If
    CC.AddDimension "yr", "Yr", xda_vertical, 1
    CC.AddDimension "mth", "Mth", xda_vertical, 2
    CC.AddDimension "wk", "Wk", xda_outside, 3
    CC.AddDimension "BIC", "BIC", xda_outside, 1
    CC.AddDimension "Br", "Br", xda_vertical, 1
    CC.AddDimension "Acno", "Acno", xda_outside, 1
    CC.AddDimension "ProductType", "ProductType", xda_outside, 1
    CC.AddDimension "EXCHANGENUMBER", "EXCHANGENUMBER", xda_outside, 1
    CC.AddDimension "COMBO", "Title", xda_outside, 1
    CC.AddDimension "P_Publisher", "Publisher", xda_outside, 1
    CC.AddDimension "P_SP", "S.P.", xda_outside, 1
    CC.AddDimension "P_Cost", "Cost", xda_outside, 1
  '  CC.AddFact "rank", "VAL", xfam_RANKA, "Rank"
    CC.AddFact "Qty", "QTY", xfaa_SUM, "Qty"
    CC.AddFact "VAL", "VAL", xfaa_SUM, "sales value"
'    CC.AddFact "VAL3", "VAL", xfam_RANKA, "Rank asc"
    CC.FieldFormat("Qty") = "##0"
    CC.FieldFormat("VAL") = "###,##0.00"
    CC.Active = False
    DoEvents
    Screen.MousePointer = vbHourglass
    CC.DataSourceType = xcdt_Recordset
'    rs.MoveFirst
'    If Not rs.EOF Then
'        CC.Open rs
'       ' CC.SortByFact xda_vertical, 1
'       ' CC.Active = True
'    Else
'        MsgBox "No records", , "Status"
'    End If
   
    Me.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub ConnectToData()
    On Error GoTo errHandler
    If rs.RecordCount < 1 Then
        MsgBox "No records", , "Status"
        Exit Sub
    End If
    rs.MoveFirst
    If Not rs.EOF Then
        CC.Open rs
    Else
        MsgBox "No records", , "Status"
    End If

    Exit Sub
errHandler:
    ErrorIn "frmProductPT.ConnectToData"
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    CC.Width = Me.Width - (CC.left + 400)
    lngDiff = CC.Height
    CC.Height = Me.Height - (CC.top + 1220)
    lngDiff = CC.Height - lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    cmdPrintCube.top = cmdPrintCube.top + lngDiff
    cmdLoad.top = cmdLoad.top + lngDiff
    cmdSaveCube.top = cmdSaveCube.top + lngDiff
    Me.cmdPrint.top = cmdPrint.top + lngDiff
    txtCubeName.top = txtCubeName.top + lngDiff
    Label1.top = Label1.top + lngDiff
    cmdConnect.top = cmdConnect.top + lngDiff
End Sub
