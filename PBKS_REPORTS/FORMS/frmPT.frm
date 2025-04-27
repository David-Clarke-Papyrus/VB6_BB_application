VERSION 5.00
Object = "{CCA2C66D-33FD-11D5-8D72-005004532BDF}#1.3#0"; "CCubeX.ocx"
Begin VB.Form frmPT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Documents per supplier"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   Begin CCubeX.ContourCubeX CC 
      Height          =   4890
      Left            =   75
      TabIndex        =   0
      Top             =   660
      Width           =   11550
      BackColor       =   14737632
      Enabled         =   -1  'True
      MainAxis        =   0
      DataSourceType  =   0
      ConnectionString=   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=PT;Data Source=PAPYRUS-94TNP9S"
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
      CubeTitle       =   "Documents per supplier per period"
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
      FIELDS_SETTINGS =   $"frmPT.frx":0000
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
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5670
      Width           =   1260
   End
End
Attribute VB_Name = "frmPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset


Public Sub Component(pRS As ADODB.Recordset, pType As String)
    Set rs = pRS
    If UCase(pType) = "SUPPLIER" Then
        Me.Caption = "Documents per supplier"
    ElseIf UCase(pType) = "STORE" Then
        Me.Caption = "Documents per store"
    End If
End Sub
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
Dim rs As New ADODB.Recordset
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

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    CC.ExportToFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", xet_html
    CC.PrintCube True, False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
top = 400
left = 20
Width = 11900
Height = 6800

    CC.AddDimension "SupplierName", "Supplier", xda_vertical, 1
    CC.AddDimension "Document", "Document", xda_vertical, 2
    CC.AddDimension "Description", "Description", xda_vertical, 3
    CC.AddDimension "DocumentType", "Document type", xda_horizontal, 1
    CC.AddDimension "PT_Code", "Product type", xda_outside, 1
    CC.AddFact "Qty", "QTY", xfaa_SUM, "Qty"
    CC.AddFact "Val", "Val", xfaa_SUM, "Value"
    CC.DimFlags("Description") = xfNoTotals + xfNoGrandTotals
    CC.DimFlags("DocumentType") = xfNoTotals + xfNoGrandTotals
'    CC.AddFormula "Turn", "DL_Value", "testTurn"
    CC.FieldFormat("Qty") = "##,##0"
    
    CC.Active = False
    
    DoEvents
    Screen.MousePointer = vbHourglass
    CC.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        CC.Open rs
        CC.Active = True
    Else
        MsgBox "No records", , "Status"
    End If
   
    Me.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.Form_Load", , EA_NORERAISE
    HandleError
End Sub
