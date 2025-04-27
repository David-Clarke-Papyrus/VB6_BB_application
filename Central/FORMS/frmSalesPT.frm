VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmSalesPT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Documents per supplier"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   11880
   Begin CCubeX2.ContourCubeX ContourCubeX 
      Height          =   4200
      Left            =   -15
      TabIndex        =   7
      Top             =   1305
      Width           =   11655
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   10528950
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   ""
      InactiveDimAreaBkColor=   10528950
      AutoSize        =   0   'False
      UnusedDataAreaColor=   -2147483643
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmSalesPT.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3690
      Top             =   5550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8100
      Top             =   5595
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
      Left            =   9075
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5565
      Width           =   1260
   End
   Begin VB.TextBox txtCubeName 
      Height          =   285
      Left            =   15
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5535
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
      Left            =   4335
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5775
      Visible         =   0   'False
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
      Height          =   300
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5835
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
      Height          =   300
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5820
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
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5565
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
      TabIndex        =   0
      Top             =   6075
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSComctlLib.ImageList HotImageList 
      Left            =   10575
      Top             =   1695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":04AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":065E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":0810
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":09C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":0B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":0D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":0ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":108A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":123C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":13EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":15A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":1746
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":18F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":1AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":1C5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   10575
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":1E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":24A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":2B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":31C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":3856
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":3EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":457A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":472C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":48DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":4A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":4C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":4DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":547A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":5B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":619E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":6830
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesPT.frx":6BCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar GridToolBar 
      Height          =   660
      Left            =   -30
      TabIndex        =   8
      Top             =   -45
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   1164
      ButtonWidth     =   820
      ButtonHeight    =   1164
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList"
      HotImageList    =   "HotImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Swap rows and columns"
            Object.ToolTipText     =   "Swap rows and columns"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Collapse rows and columns"
            Object.ToolTipText     =   "Collapse rows and columns"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Expand rows and columns"
            Object.ToolTipText     =   "Expand rows and columns"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Percents by rows|Calculate percents in rows and show its in cells"
            Object.ToolTipText     =   "Percents by rows|Calculate percents in rows and show its in cells"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sort rows by fact|Sort rows by selected fact values in selected column"
            Object.ToolTipText     =   "Sort rows by fact|Sort rows by selected fact values in selected column"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "asc"
                  Object.Tag             =   "asc"
                  Text            =   "Ascending"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "desc"
                  Object.Tag             =   "desc"
                  Text            =   "Descending"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nosort"
                  Object.Tag             =   "Nosort"
                  Text            =   "No sorting"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sort columns by fact|Sort columns by selected fact values in selected row"
            Object.ToolTipText     =   "Sort columns by fact|Sort columns by selected fact values in selected row"
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hasc"
                  Object.Tag             =   "hasc"
                  Text            =   "ascending"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hdesc"
                  Object.Tag             =   "hdesc"
                  Text            =   "descending"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hNosort"
                  Object.Tag             =   "hNosort"
                  Text            =   "No sort"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Scale data"
            Object.ToolTipText     =   "Scale data"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x1"
                  Text            =   "Scale 1x1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x10"
                  Text            =   "Scale 1x10"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x100"
                  Text            =   "Scale 1x100"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x1000"
                  Text            =   "Scale 1x1'000"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Description     =   "Export with Chart|Export Grid with\without Chart"
            Object.ToolTipText     =   "Export with Chart|Export Grid with\without Chart"
            ImageIndex      =   15
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to HTML| Export Grid and Chart1 to HTML for printing and publishing"
            Object.ToolTipText     =   "Export to HTML| Export Grid and Chart1 to HTML for printing and publishing"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to Excel|Export Grid and Chart1 to Excel for printing, additioanal calculation and publishing"
            Object.ToolTipText     =   "Export to Excel|Export Grid and Chart1 to Excel for printing, additioanal calculation and publishing"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to Word|Export Grid and Chart1 to Word for printing and publishing"
            Object.ToolTipText     =   "Export to Word|Export Grid and Chart1 to Word for printing and publishing"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Print|Print Grid"
            Object.ToolTipText     =   "Print Grid"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "load"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "save"
            ImageIndex      =   17
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSalesPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset
Const opTRANSPOSE = 1
Const opCOLLAPSE = 2
Const opEXPAND = 3
Const opPERCENT = 4
Const opSORT_COL = 6
Const opSORT_ROW = 7
Const opEXPORT_HTML = 11
Const opEXPORT_XLS = 12
Const opEXPORT_DOC = 13
Const opPRINT = 15
Const opLOADLAYOUT = 16
Const opSAVELAyoUT = 17

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Sub Component(pRS As ADODB.Recordset)
    On Error GoTo errHandler
    Set rs = pRS
    Caption = "Sales patterns"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.Component(pRS)", pRS
End Sub

''''''Private Sub CC_KeyUp(ByVal KeyCode As Long, ByVal Shift As Long)
''''''    On Error GoTo errHandler
''''''    If KeyCode = vbKeyC Then
''''''            If Shift = 1 Then
''''''                CC.ViewFlags = CC.ViewFlags + xfDescending
''''''            Else
''''''                If CC.ViewFlags = xfDescending Then
''''''                    CC.ViewFlags = CC.ViewFlags - xfDescending
''''''                End If
''''''            End If
''''''            CC.SortByFact xda_vertical
''''''    End If
''''''    If KeyCode = vbKeyX Then
''''''            CC.CancelFactSorting xda_vertical
''''''
''''''            CC.DimFlags("Acno") = 0
''''''            CC.ViewFlags = 0
''''''    End If
''''''    Exit Sub
''''''errHandler:
''''''    If ErrMustStop Then Debug.Assert False: Resume
''''''    ErrorIn "frmSalesPT.CC_KeyUp(KeyCode,Shift)", Array(KeyCode, Shift)
''''''End Sub
''''''
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
    ContourCubeX.Active = False
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPT.cmdClose_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdClose_Click"
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim SQL As String

    ContourCubeX.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    WaitMsg "Loading the pivot table . . . ", True, Me
    ContourCubeX.DataSourceType = xcdt_Recordset
    ContourCubeX.Open rs
   
    ContourCubeX.Active = True
    WaitMsg "", False, Me
    Me.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdFetch_Click"
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo errHandler
    ConnectToData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdConnect_Click"
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
    CD1.DefaultExt = ".txt"
    CD1.DialogTitle = "Load stored cube"
    CD1.InitDir = "C:\PBKS\BU"
    
    CD1.ShowOpen
    txtCubeName = CD1.FileName
    If fs.FileExists(txtCubeName) Then
        ContourCubeX.Cube.Load txtCubeName
    Else
        MsgBox "Nothing to load"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdLoad_Click"
End Sub

Private Sub LoadFormat()
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Load Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
  CommonDialog1.CancelError = True
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err.Number = cdlCancel Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    LoadContourcubeLayout CommonDialog1.FileName
  End If

End Sub
Private Sub SaveFormat()
Dim fs As New FileSystemObject
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\CubeFormats") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\CubeFormats")
    End If
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Save Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
  CommonDialog1.CancelError = True
  On Error Resume Next
  CommonDialog1.ShowSave
  If Err.Number = cdlCancel Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    If Trim(CommonDialog1.FileName) <> "" Then SaveContourCubeLayout CStr(CommonDialog1.FileName)
  End If

End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim res As Boolean
Dim fs As New FileSystemObject

    ContourCubeX.ExportToFile oPC.SharedFolderRoot & "\HTML\SalesPatterns.html", oPC.SharedFolderRoot & "\HTML\SalesPatterns.html", xet_html
    MsgBox "Exported to file " & oPC.SharedFolderRoot & "\HTML\SalesPatterns.html"
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdPrint_Click"
End Sub

Private Sub cmdPrintCube_Click()
    On Error GoTo errHandler
    ContourCubeX.AllowTitle = True
    ContourCubeX.CubeTitle = "Sales patterns: printed " & Format(Now(), "DD/mm/yyyy HH:HH AM/PM")
  '  CC.CubeFooter = "TEST Footer"

    ContourCubeX.PrintCube True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdPrintCube_Click"
End Sub

Private Sub cmdSaveCube_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject

    CD1.DefaultExt = ".txt"
    CD1.DialogTitle = "Save cube"
    CD1.InitDir = "C:\PBKS\BU"
    
  On Error Resume Next
    CD1.ShowOpen
  If CD1.FileName = "" Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    txtCubeName = CD1.FileName
    If fs.FileExists(txtCubeName) Then
        fs.DeleteFile txtCubeName
    End If
    ContourCubeX.Cube.Save txtCubeName
  End If
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.cmdSaveCube_Click"
End Sub



Private Sub ContourCubeX_OnCubeLoaded()
    On Error GoTo errHandler
    CheckVisible
    CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.ContourCubeX_OnCubeLoaded"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSImple
    top = 400
    left = 20
    Width = 11900
    Height = 6800

    
    DoEvents
   
    Me.Refresh
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.Form_Load"
End Sub
Private Sub ConnectToData()
    On Error GoTo errHandler
  Dim Fact As IViewFact
    If rs.RecordCount < 1 Then
        MsgBox "No records", , "Status"
        Exit Sub
    End If
    rs.MoveFirst
    If Not rs.EOF Then
'        ContourCubeX.Cube.Open rs, True
    Else
        MsgBox "No records", , "Status"
    End If
    If Not rs.EOF Then
 CloseCube
  
    With ContourCubeX.Cube
        .Dims.Add("Rank", "Rank_Rank", , xda_vertical).MoveTo xda_vertical, 0
        .Dims.Add("yr", "Yr", , xda_vertical).MoveTo xda_outside
        .Dims.Add("Br", "Br", , xda_vertical).MoveTo xda_vertical, 1
        
        .Dims.Add("mth", "Mth", , xda_vertical).MoveTo xda_outside
        .Dims.Add("wk", "Wk", , xda_vertical).MoveTo xda_outside
        .Dims.Add("COMBO", "Title", , xda_vertical).MoveTo xda_outside
        .Dims.Add("Cost", "P_Cost", , xda_vertical).MoveTo xda_outside
        .Dims.Add("SP", "P_SP", , xda_vertical).MoveTo xda_outside
        .Dims.Add("AcnoCustomer", "Acno", , xda_vertical).MoveTo xda_outside, 1
        .Dims.Add("Acno", "CustomerAcno", , xda_vertical).MoveTo xda_outside, 1
        .Dims.Add("EXCHANGENUMBER", "EXCHANGENUMBER", , xda_vertical).MoveTo xda_outside, 1
        .Dims.Add("HBr", "STORE_CODE", , xda_vertical).MoveTo xda_outside, 1
        
        .BaseFacts.Add "bfQTY", "QTY"
        .BaseFacts.Add "bfVAL", "VAL"
        .Facts.Add "QTY", "bfQTY", xfaa_SUM
        .Facts.Add "VAL", "bfVAL", xfaa_SUM
        ContourCubeX.Facts(0).Appearance.Format = "##0"
        ContourCubeX.Facts(1).Appearance.Format = "###,##0.00"
        For Each Fact In ContourCubeX.Facts
          Fact.Visible = True
        Next
        Set rs.ActiveConnection = Nothing
        '.DataSourceType = xcdt_Recordset
        .Open rs
    End With
  
    AfterOpen
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.ConnectToData"
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
 ContourCubeX.Visible = ContourCubeX.Active
 CheckEnabled
 CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.AfterOpen"
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
 CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.Form_Activate"
End Sub

Private Sub CloseCube()
    On Error GoTo errHandler
 With ContourCubeX
   .Active = False
   .Cube.Dims.Clear
   .Cube.Facts.Clear
   .Cube.BaseFacts.Clear
 End With
 CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CloseCube"
End Sub
Private Sub CheckEnabled()
    On Error GoTo errHandler
 Dim i As Integer
 'Check if controls are enabled or not
 For i = 1 To GridToolBar.Buttons.Count
  GridToolBar.Buttons(i).Enabled = ContourCubeX.Active
 Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckEnabled"
End Sub

Private Sub CheckVisible()
    On Error GoTo errHandler
 ContourCubeX.Visible = True 'ContourCubeX.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckVisible"
End Sub
Private Sub GridToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim DDLevel As Integer
 Dim Checked As Boolean
  
 Checked = (Button.Value = tbrPressed)
        ContourCubeX.TitleSettings.Text = "TEST"

 With ContourCubeX
  Select Case Button.Index
   Case opTRANSPOSE          'Swap rows and columns
    .Transposed = Checked
    .Cube.RootAxis = IIf(.Transposed, _
     IIf(GridToolBar.Buttons(6).Value = tbrPressed, xda_vertical, xda_horizontal), _
     IIf(GridToolBar.Buttons(6).Value = tbrPressed, xda_horizontal, xda_vertical))
   Case opCOLLAPSE           'Expand/Collapse rows and columns
    If .HAxis.Dims.Count > 0 Then .HAxis.DrillDownLevel = 0
    If .VAxis.Dims.Count > 0 Then .VAxis.DrillDownLevel = 0
   Case opEXPAND
    .HAxis.DrillDownLevel = .HAxis.Width - 1
    .VAxis.DrillDownLevel = .VAxis.Width - 1
   Case opPERCENT            'Calculate percents by rows/columns and show it in cells
    .Active = False
    Dim Fact As ICubeFact
    For Each Fact In .Cube.Facts
      If left(Fact.Name, 3) <> "_P_" Then
        If Not .Cube.Facts.Exists("_P_" & Fact.Name) Then
          .Cube.Facts.AddFormula("_P_" & Fact.Name, Fact.Name & "/%Total(" & Fact.Name & ")").Active = True
        End If
        If GridToolBar.Buttons(4).Value = tbrPressed Then
          .Facts.Item("_P_" & Fact.Name).Visible = True
          .Facts.Item("_P_" & Fact.Name).Caption = Fact.Caption
          .Facts.Item("_P_" & Fact.Name).Appearance.Format = "#####0.00%"
          .Facts.Item(Fact.Name).Enabled = False
        Else
          .Facts.Item(Fact.Name).Visible = True
          .Facts.Item("_P_" & Fact.Name).Enabled = False
        End If
      End If
    Next
    .Active = True

   Case opSORT_COL, opSORT_ROW        'Sort rows by selected fact values in selected column
    Dim SortAxis: SortAxis = IIf(Button.Index = 6, xda_vertical, xda_horizontal)
    Dim Col As Long, Row As Long: Col = .CurrentCell.Col: Row = .CurrentCell.Row
    If (GridToolBar.Buttons(Button.Index).Value = tbrPressed) Then _
      .SortGridByFact SortAxis, Col, Row _
    Else _
      .CancelFactSorting (SortAxis)
   'Export Grid for printing and publishing
   Case opEXPORT_HTML
    ExportCube .TitleSettings.Text, xolaprpt_HTML, "html"
   Case opEXPORT_XLS
    ExportCube .TitleSettings.Text, xolaprpt_XLS, "xls"
   Case opEXPORT_DOC
    ExportCube .TitleSettings.Text, xolaprpt_HTML, "doc"
   Case opPRINT
    .PrintCube True, False
   Case opSAVELAyoUT
        SaveFormat
   Case opLOADLAYOUT
        LoadFormat
  End Select
 End With
End Sub

Private Sub GridToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Dim ScaleFactor As Double
 Dim SortAxis
 Dim Col As Long, Row As Long
 ScaleFactor = 1
 With ContourCubeX
  Select Case ButtonMenu.Key
   Case "1x1"
    ScaleFactor = 1
   Case "1x10"
    ScaleFactor = 0.1
   Case "1x100"
    ScaleFactor = 0.01
   Case "1x1000"
    ScaleFactor = 0.001
   Case "asc"
       SortAxis = xda_vertical
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            ContourCubeX.Descending = False
            ContourCubeX.SortGridByFact SortAxis, Col, Row
   Case "desc"
       SortAxis = xda_vertical
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            ContourCubeX.Descending = True
            ContourCubeX.SortGridByFact SortAxis, Col, Row
   Case "Nosort"
       SortAxis = xda_vertical
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            ContourCubeX.CancelFactSorting (SortAxis)
   Case "hasc"
       SortAxis = xda_horizontal
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            ContourCubeX.Descending = False
            ContourCubeX.SortGridByFact SortAxis, Col, Row
   Case "hdesc"
       SortAxis = xda_horizontal
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            ContourCubeX.Descending = True
            ContourCubeX.SortGridByFact SortAxis, Col, Row
   Case "hNosort"
       SortAxis = xda_horizontal
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            ContourCubeX.CancelFactSorting (SortAxis)
  End Select
  Dim Fact
  For Each Fact In .Facts
   If Fact.Enabled Then Fact.ScaleFactor = ScaleFactor
  Next
 End With
End Sub
Private Sub ExportCube(FileName As String, FileFormat As TxOlapReportType, FileType As String)
 'Export OLAP-report to Excel, Word, HTML as file in html format
 FileName = FileName + "." + FileType
 ContourCubeX.ReportToFile FileName, "", FileFormat
 OpenDocument (FileName)
End Sub
Private Sub OpenDocument(f_name As String)
 Dim Scr_hDC As Long
 Scr_hDC = GetDesktopWindow()
 ShellExecute Scr_hDC, "Open", f_name, "", "", 1
End Sub


Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    ContourCubeX.Width = Me.Width - (ContourCubeX.left + 400)
    lngDiff = ContourCubeX.Height
    ContourCubeX.Height = Me.Height - (ContourCubeX.top + 1220)
    lngDiff = ContourCubeX.Height - lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    cmdPrintCube.top = cmdPrintCube.top + lngDiff
    cmdLoad.top = cmdLoad.top + lngDiff
    cmdSaveCube.top = cmdSaveCube.top + lngDiff
    Me.cmdPrint.top = cmdPrint.top + lngDiff
    txtCubeName.top = txtCubeName.top + lngDiff
  '  Label1.top = Label1.top + lngDiff
    cmdConnect.top = cmdConnect.top + lngDiff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.Form_Resize"
End Sub
Public Sub SaveContourCubeLayout(ltFile As String)
'Saving layout procedure
  Dim rsFields, Axis, Object, bInvertFilterSelection, Value, i, j, viewTotalsState, _
      viewGTotalsState, strExpand, fs
  rsFields = Array("Object", "Name", "Property", "Value")
  'Create an ADO recordset with 4 fields:
  Dim rs As New ADODB.Recordset
  rs.Fields.Append rsFields(0), adBSTR, 10
  rs.Fields.Append rsFields(1), adBSTR, 50
  rs.Fields.Append rsFields(2), adVariant, 50
  rs.Fields.Append rsFields(3), adVariant, 255
  rs.Open
  rs.AddNew rsFields, Array("Cube", ContourCubeX.Name, "RootAxis", ContourCubeX.Cube.RootAxis)
  With ContourCubeX
    'Populate recordset with layout properties
    For Each Object In .Facts
      'Fact visibility
      rs.AddNew rsFields, Array("Fact", Object.Name, "Visible", Object.Visible)
    Next
    For i = 0 To 1
        If i = 0 Then Set Axis = .VAxis Else Set Axis = .HAxis
        For Each Object In Axis.Dims
          'Dimension positions and properties
          rs.AddNew rsFields, Array("Dim", Object.Name, "Axis", Object.CubeDim.Axis)
          rs.AddNew rsFields, Array("Dim", Object.Name, "Pos", Object.CubeDim.pos)
        Next
    Next
    For Each Object In .Dims
        rs.AddNew rsFields, Array("Dim", Object.Name, "Totals", Object.NoTotals)
        rs.AddNew rsFields, Array("Dim", Object.Name, "Descending", Object.Descending)
        'Dimension filters:
        'To minimize the file, choose the minimum set between hidden and visible
        'values to save
        bInvertFilterSelection = (Object.CubeDim.GetValues(2).Count > Object.CubeDim.GetValues(1).Count)
        rs.AddNew rsFields, Array("DimsFilter", "InvertFilterSelection", Object.Name, bInvertFilterSelection)
        For Each Value In Object.CubeDim.GetValues(IIf(bInvertFilterSelection, 1, 2))
          rs.AddNew rsFields, Array("DimsFilter", "Filter", Object.Name, Value)
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
        rs.AddNew rsFields, Array("Axis", "Horizontal", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    If .VAxis.Length > 0 Then
      For i = 0 To .VAxis.Length - 1
        strExpand = ""
        For j = 0 To .VAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .VAxis.GetSection(i).getValue(j)
        Next j
        rs.AddNew rsFields, Array("Axis", "Vertical", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    'Restore view totals
    .NoTotals = viewTotalsState
    .NoGrandTotals = viewGTotalsState
  End With
  'Verify if the file already exists and eventually delete it before saving
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.FileExists(ltFile) Then fs.DeleteFile (ltFile)
  rs.Save ltFile, adPersistXML
  rs.Close
End Sub

Sub LoadContourcubeLayout(ltFile As String)
'Loading layout procedure
  Dim FactSettings, DimSettings, Object, DimFilters, AxisSettings, i, bInvertFilterSelection
  Dim rs As New ADODB.Recordset
  'First open the saved XML layout file
  rs.Open ltFile
  With ContourCubeX
    'Restore cube properties
    rs.Filter = "Object='Cube'"
    .Cube.RootAxis = CInt(rs.GetRows()(3, 0))
    'Fact visibility
    rs.Filter = "Object='Fact'"
    FactSettings = rs.GetRows()
    For i = 0 To UBound(FactSettings, 2)
      If LCase(CStr(FactSettings(2, i))) = "visible" Then
        If .Facts.Exists(CStr(FactSettings(1, i))) Then _
           .Facts(CStr(FactSettings(1, i))).Visible = CBool(FactSettings(3, i))
      End If
    Next i
    'Set up dimension positions, totalling and sort orders
    rs.Filter = "Object='Dim'"
    DimSettings = rs.GetRows()
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
    .Active = True
    'Dimension filter states
    rs.Filter = "Object='DimsFilter'"
    DimFilters = rs.GetRows()
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
    rs.Filter = "Object='Axis'"
    AxisSettings = rs.GetRows()
    For i = 0 To UBound(AxisSettings, 2)
      ExpandSection CStr(AxisSettings(1, i)), CStr(AxisSettings(3, i))
    Next i
  End With
  rs.Close
End Sub

Sub ExpandSection(strAxis As String, strExpand As String)
'This procedure restores saved state of an axis section
'It searches for given combination of dim values along the axis,
'and expands the section found
  Dim Axis As IViewAxis, i, j, aExpand
  aExpand = Split(strExpand, Chr(10))
  If LCase(strAxis) = "horizontal" Then Set Axis = ContourCubeX.HAxis Else Set Axis = ContourCubeX.VAxis
  On Error Resume Next
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
  On Error GoTo 0
End Sub
