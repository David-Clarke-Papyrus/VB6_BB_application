VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmTradingPerformance 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Trading performance"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13335
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5490
      Left            =   120
      TabIndex        =   4
      Top             =   795
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   9684
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "By publisher"
      TabPicture(0)   =   "frmTradingPerformance3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHeading"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboDisplay_Publisher"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGO_Publisher"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ARViewer21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "By category"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin CCubeX2.ContourCubeX CC 
         Height          =   2535
         Left            =   225
         TabIndex        =   8
         Top             =   750
         Width           =   12555
         Active          =   0   'False
         Transposed      =   0   'False
         NULLValueString =   ""
         Descending      =   0   'False
         NoTotals        =   0   'False
         NoGrandTotals   =   0   'False
         Caption         =   ""
         BackColor       =   13882315
         Enabled         =   -1  'True
         Alive           =   0   'False
         BorderStyle     =   1
         AllowInactiveDimArea=   -1  'True
         AllowExpand     =   -1  'True
         AllowPivot      =   -1  'True
         TotalsString    =   ""
         InactiveDimAreaBkColor=   13160660
         AutoSize        =   0   'False
         UnusedDataAreaColor=   -2147483643
         MousePointer    =   0
         Object.Visible         =   -1  'True
         InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
         ConnectionString=   ""
         DataSourceType  =   0
         VERSION_NO      =   2
         CCubeXMetadata  =   $"frmTradingPerformance3.frx":001C
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
         Height          =   1875
         Left            =   435
         TabIndex        =   11
         Top             =   3360
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   3307
         SectionData     =   "frmTradingPerformance3.frx":045A
      End
      Begin VB.CommandButton cmdGO_Publisher 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   12165
         Picture         =   "frmTradingPerformance3.frx":0496
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   375
         Width           =   600
      End
      Begin VB.ComboBox cboDisplay_Publisher 
         Height          =   315
         ItemData        =   "frmTradingPerformance3.frx":0820
         Left            =   8610
         List            =   "frmTradingPerformance3.frx":0851
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Sales"
         Top             =   390
         Width           =   3420
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   " (retail inc VAT)"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   390
         Width           =   8130
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54394881
      CurrentDate     =   37421
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   7470
      Top             =   105
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
            Picture         =   "frmTradingPerformance3.frx":098F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":1021
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":16B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":1D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":23D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":2A69
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":30FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":32AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":345F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":3611
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":37C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":3969
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":3FFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":468D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":4D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":53B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTradingPerformance3.frx":574B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar GridToolBar 
      Height          =   660
      Left            =   75
      TabIndex        =   9
      Top             =   45
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1164
      ButtonWidth     =   820
      ButtonHeight    =   1164
      Wrappable       =   0   'False
      Appearance      =   1
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This report set is still under construction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   435
      Left            =   4830
      TabIndex        =   10
      Top             =   6375
      Width           =   4380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Since"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   8220
      TabIndex        =   1
      Top             =   285
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmTradingPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As adodb.Recordset

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



''''Private Sub CC_OnDimValueClick(ByVal DimName As String, ByVal axis As CCubeX.TxDimAxis, ByVal ValNo As Long, ByVal Level As Long, ByVal Value As String, ByVal Totals As Boolean)
''''On Error Resume Next
''''    CC.SortByFact xda_vertical, ValNo, 1
''''    CC.DimFlags("SupplierName") = CC.DimFlags("SupplierName") + xfDescending
''''
''''End Sub
''''
''''Private Sub CC_OnFactHeadingClick(ByVal FactName As String, ByVal Col As Long)
''''    CC.SortByFact xda_vertical, Col, 1
''''End Sub
''''
''''Private Sub CC_OnFactHeadingDblClick(ByVal FactName As String, ByVal Col As Long)
''''    CC.SortByFact xda_vertical, Col, 1
''''End Sub
''''
''''Private Sub CC_OnTitleClick()
''''   ' CC.SortByFact xda_vertical, Col, 1
''''End Sub

Private Sub cmdGO_Publisher_Click()
   ' me.cboDisplay_Publisher.Locked
    Preparecube Me.cboDisplay_Publisher
End Sub
Private Sub CC_OnFactValueDblClick(ByVal FactName As String, ByVal Col As Long, ByVal Row As Long, ByVal Value As String)
Dim v As CCubeX.IContourViewX
Dim SuppID As Long
Dim Mth As Date
Dim lRS As adodb.Recordset
Dim f As frmPerformance
Dim x As Long
Dim Y As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set f = New frmPerformance

    x = MouseX(Me.hwnd)
    Y = MouseY(Me.hwnd)
    If Me.Width / Screen.TwipsPerPixelX / 2 > x Then
        x = x + 11
    Else
        x = x - (f.Width / Screen.TwipsPerPixelX) - 30 '- 200
    End If
    If Me.Height / Screen.TwipsPerPixelY / 2 > Y Then
        Y = Y + 11
    Else
        Y = Y - (f.Height / Screen.TwipsPerPixelY) '- 30 '- 200
    End If
    CC.GetView -1, -1, v
    Set lRS = New adodb.Recordset
    lRS.CursorLocation = adUseClient
    lRS.Open "SELECT " _
        & "TP_NAME as supplierName, " _
        & "PERF_Month as ReportMonth, " _
        & "PERF_StockValue_CostEx, " _
        & "PERF_SalesValue_CostEx, " _
        & "PERF_SalesValue_RetailInc, " _
        & "PERF_SalesValue_RetailInc / 1.14 as PERF_SalesValue_RetailEx, " _
        & "PERF_SalesAsPercentOfTotalSales_RetailInc, " _
        & "PERF_SalesAsPercentOfTotalSOH_RetailInc, " _
        & "PERF_StockAsPercentOfTotalSOH_CostEx, " _
        & "CASE WHEN ISNULL(PERF_StockValue_RetailInc,0) = 0 OR ISNULL(SUPPS_QtyMonthsInStockTurnRange,0) = 0 THEN 0 ELSE SUPPS_Last12MonthSalesValue*(12/SUPPS_QtyMonthsInStockTurnRange)/(SUPPS_Last12MonthStockValue/SUPPS_QtyMonthsInStockTurnRange) END as StockTurn, " _
        & "PERF_ReturnsASPercentDeliveries, " _
        & "PERF_ReturnsASPercentSales, " _
        & "PERF_Margin, " _
        & "PERF_OrdersPlacedValue_CostEx, " _
        & "PERF_OrdersOSValue_CostEx, " _
        & "PERF_MissinglastStockTake_CostEx " _
        & " FROM tPERFORMANCE a JOIN tTP b ON a.PERF_SUPPLIERID = b.TP_ID JOIN tSupplierStatsMonthly ON PERF_MONTH = SUPPS_MONTH AND PERF_SUPPLIERID = SUPPS_SUPPLIERID WHERE PERF_Month = CAST('" & ReverseDate(v.GetDimValue(1, Col, xctLeaf)) & "' AS DATETIME) AND TP_NAME = '" & Replace(v.GetDimValue(0, Row, xctLeaf), "'", "''") & "'", oPC.COSHORT
    If lRS.RecordCount = 0 Then Exit Sub
    f.Component lRS
    f.top = Y * Screen.TwipsPerPixelY
    f.left = x * Screen.TwipsPerPixelX
    f.Show vbModal
    lRS.Close
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



Private Sub cmdOK_Click()
Dim oSQL As New z_SQL
  MsgBox "Code skipped"
    'oSQL.RunProc "CreatePerformanceTable", Array(), "Loading performance table"
    
    Set rs = oSQL.GetPerformanceStats
    
    Preparecube Me.cboDisplay_Publisher
 '  Me.lblHeading.Caption = "NOTE: All retail values include VAT, cost values exclude VAT"
    Me.cboDisplay_Publisher.Enabled = True
    MsgBox "Loaded"
End Sub

Private Sub Preparecube(Style As String)
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    cboDisplay_Publisher.Locked = False
    If rs Is Nothing Then Exit Sub
    rs.MoveFirst
    If rs.EOF Then
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.EOF Then
        
        CloseCube
        
    '    CC.Active = False
    '    CC.ClearFields
        With CC.Cube
            .Dims.Add("Supplier", "SupplierName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Mth", "PERF_Month", , xda_vertical).MoveTo xda_horizontal
    '    CC.AddDimension "SupplierName", "Supplier", xda_vertical, 1
    '    CC.AddDimension "PERF_Month", "Mth", xda_horizontal, 1
            If Style = "QtyInTop50" Then
                .BaseFacts.Add "QtyInTop50", "PERF_QtyInTop50"
                .Facts.Add "PERF_QtyInTop50", "QtyInTop50", xfaa_SUM
                CC.Facts(0).Appearance.Format = "##0"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
               ' CC.AddFact "QtyInTop50", "PERF_QtyInTop50", xfaa_SUM, "QtyInTop50"
               ' CC.FieldFormat("QtyInTop50") = "###,###.00;-###,###;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
                CC.TitleSettings.Text = "Number of titles in top 50"
            ElseIf Style = "Sales" Then
                .BaseFacts.Add "SalesValue", "PERF_Salesvalue_RetailInc"
                .Facts.Add "PERF_Salesvalue_RetailInc", "SalesValue", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Sales"
        
        '        CC.AddFact "SalesValue", "PERF_Salesvalue_RetailInc", xfaa_SUM, "PERF_StockValue_RetailInc"
        '        CC.FieldFormat("SalesValue") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.CubeTitle = "Sales"
            ElseIf Style = "Sales as percent of Total sales" Then
                .BaseFacts.Add "PercentOfSales", "PERF_SalesAsPercentOfTotalSales_RetailInc"
                .Facts.Add "PERF_SalesAsPercentOfTotalSales_RetailInc", "PercentOfSales", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Sales as percent of Total sales"
        '        CC.AddFact "PercentOfSales", "PERF_SalesAsPercentOfTotalSales_RetailInc", xfaa_SUM, "Percent of sales"
        '        CC.FieldFormat("PercentOfSales") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.CubeTitle = "Sales as percent of Total sales"
            ElseIf Style = "Sales as percent of Total stock" Then
                .BaseFacts.Add "PercentOfSales", "PERF_SalesAsPercentOfTotalSOH_RetailInc"
                .Facts.Add "PERF_SalesAsPercentOfTotalSOH_RetailInc", "PercentOfSales", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Sales as percent of Total stock"
        '        CC.AddFact "PercentOfSales", "PERF_SalesAsPercentOfTotalSOH_RetailInc", xfaa_SUM, "Percent of sales"
        '        CC.FieldFormat("PercentOfSales") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.CubeTitle = "Sales as percent of Total stock"
            ElseIf Style = "Stock value" Then
                .BaseFacts.Add "StockValue", "PERF_StockValue_CostEx"
                .Facts.Add "PERF_StockValue_CostEx", "StockValue", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Stock value (at cost Ex VAT)"
        '        CC.AddFact "StockValue", "PERF_StockValue_CostEx", xfaa_SUM, "PERF_StockValue"
        '        CC.FieldFormat("StockValue") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfNoTotals + xfNoGrandTotals
        '        CC.CubeTitle = "Stock value (at cost Ex VAT)"
            ElseIf Style = "Stock value as percent of total stock" Then
                .BaseFacts.Add "PercentOfSales", "PERF_StockAsPercentOfTotalSOH_CostEx"
                .Facts.Add "PERF_StockAsPercentOfTotalSOH_CostEx", "PercentOfSales", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Stock turn"
        '        CC.AddFact "PercentOfSales", "PERF_StockAsPercentOfTotalSOH_CostEx", xfaa_SUM, "Percent of sales"
        '        CC.FieldFormat("PercentOfSales") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfNoTotals + xfNoGrandTotals
        '        CC.CubeTitle = "Stock value as percent of total stock"
            ElseIf Style = "Stock turn" Then
                .BaseFacts.Add "StockTurn", "StockTurn"
                .Facts.Add "StockTurn", "StockTurn", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Stock turn"
        '        CC.AddFact "StockTurn", "StockTurn", xfaa_SUM, "StockTurn"
        '        CC.FieldFormat("StockTurn") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero + xfDescending
        '        CC.CubeTitle = "Stock turn"
            ElseIf Style = "Returns" Then
                .BaseFacts.Add "Returns", "PERF_Returns_RetailInc"
                .Facts.Add "PERF_Returns_RetailInc", "Returns", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Returns"
        '        CC.AddFact "Returns", "PERF_Returns_RetailInc", xfaa_SUM, "Returns (Cost ex VAT)"
        '        CC.FieldFormat("Returns") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Returns"
            ElseIf Style = "Returns as percent of deliveries" Then
                .BaseFacts.Add "PercentOfDeliveries", "PERF_ReturnsAsPercentDeliveries"
                .Facts.Add "PERF_ReturnsAsPercentDeliveries", "PercentOfDeliveries", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Returns as percent of deliveries"
        '        CC.AddFact "PercentOfDeliveries", "PERF_ReturnsAsPercentDeliveries", xfaa_SUM, "Percent of deliveries"
        '        CC.FieldFormat("PercentOfDeliveries") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Returns as percent of deliveries"
            ElseIf Style = "Returns as percent of sales" Then
                .BaseFacts.Add "ReturnsPercentOfSales", "PERF_ReturnsAsPercentSales"
                .Facts.Add "PERF_ReturnsAsPercentSales", "ReturnsPercentOfSales", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Returns as percent of sales (calculated at cost values)"
        '        CC.AddFact "ReturnsPercentOfSales", "PERF_ReturnsAsPercentSales", xfaa_SUM, "Returns percent of sales"
        '        CC.FieldFormat("ReturnsPercentOfSales") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Returns as percent of sales (calculated at cost values)"
            ElseIf Style = "Margin" Then
                .BaseFacts.Add "Margin", "PERF_Margin"
                .Facts.Add "PERF_Margin", "Margin", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Margin"
        '        CC.AddFact "Margin", "Margin", xfaa_SUM, "Margin"
        '        CC.FieldFormat("Margin") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Margin"
            ElseIf Style = "Margin percentage" Then
                .BaseFacts.Add "Margin percentage", "PERF_Margin"
                .Facts.Add "PERF_Margin", "Margin percentage", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Margin percentage"
        '        CC.AddFact "Margin percentage", "PERF_Margin", xfaa_SUM, "Margin percentage"
        '        CC.FieldFormat("Margin percentage") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Margin percentage"
            ElseIf Style = "Orders placed" Then
                .BaseFacts.Add "Orders placed", "PERF_OrdersPlacedValue_CostEx"
                .Facts.Add "PERF_OrdersPlacedValue_CostEx", "Orders placed", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Orders placed (cost ex VAT)"
        '        CC.AddFact "Orders placed", "PERF_OrdersPlacedValue_CostEx", xfaa_SUM, "Orders placed"
        '        CC.FieldFormat("Orders placed") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Orders placed (cost ex VAT)"
            ElseIf Style = "Orders outstanding" Then
                .BaseFacts.Add "Orders outstanding", "PERF_OrdersOSValue_CostEx"
                .Facts.Add "PERF_OrdersOSValue_CostEx", "Orders outstanding", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Orders outstanding (cost ex VAT)"
        '        CC.AddFact "Orders outstanding", "PERF_OrdersOSValue_CostEx", xfaa_SUM, "Orders outstanding"
        '        CC.FieldFormat("Orders outstanding") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Orders outstanding (cost ex VAT)"
            ElseIf Style = "Missing value last stock take" Then
                .BaseFacts.Add "Missing", "PERF_MissingLastStockTake_RetailInc"
                .Facts.Add "PERF_MissingLastStockTake_RetailInc", "Missing", xfaa_SUM
                CC.Facts(0).Appearance.Format = "###,###.00;-###,###.00;###"
                CC.NoGrandTotals = True
                CC.Dims(0).NoTotals = True
                CC.Dims(1).NoTotals = True
                CC.TitleSettings.Text = "Missing value last stock take (Retail inc VAT)"
        '        CC.AddFact "Missing", "PERF_MissingLastStockTake_RetailInc", xfaa_SUM, "Missing"
        '        CC.FieldFormat("Missing") = "###,###.00;-###,###.00;###"
        '        CC.DimFlags("PERF_Month") = xfExcludeZero + xfNoTotals + xfNoGrandTotals
        '        CC.DimFlags("SupplierName") = xfNoTotals + xfNoGrandTotals + xfExcludeZero
        '        CC.CubeTitle = "Missing value last stock take (Retail inc VAT)"
            End If
        
        
            For Each Fact In CC.Facts
              Fact.Visible = True
            Next
            Set rs.ActiveConnection = Nothing
            .Open rs

        End With
        AfterOpen
'        CC.AllowTitle = True
'        CC.TitleBkColor = RGB(240, 240, 240)
'        CC.TitleForeColor = vbBlue
'        CC.SuppressZeroCols = True
'        CC.SuppressZeroRows = True
'
'        CC.Active = False
'            CC.SuppressZeroCols = True
'            CC.SuppressZeroRows = True
'        DoEvents
'        Screen.MousePointer = vbHourglass
'        CC.DataSourceType = xcdt_Recordset
'        rs.MoveFirst
'        If Not rs.EOF Then
'            CC.Open rs
'            CC.Active = True
'            CC.SuppressZeroCols = True
'            CC.SuppressZeroRows = True
'        Else
'            MsgBox "No records", , "Status"
'        End If
'
'        Me.Refresh
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformance.Preparecube"
    HandleError
End Sub
Private Sub CloseCube()
    On Error GoTo errHandler
 With CC
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
Private Sub AfterOpen()
    On Error GoTo errHandler
 CC.Visible = CC.Active
 CheckEnabled
 CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.AfterOpen"
End Sub
Private Sub CheckEnabled()
    On Error GoTo errHandler
 Dim i As Integer
 'Check if controls are enabled or not
 For i = 1 To GridToolBar.Buttons.Count
  GridToolBar.Buttons(i).Enabled = CC.Active
 Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckEnabled"
End Sub

Private Sub CheckVisible()
    On Error GoTo errHandler
 CC.Visible = True 'ContourCubeX.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckVisible"
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next

    lngDiff = SSTab1.Height
    SSTab1.Height = Me.Height - (SSTab1.top + 700)
    lngDiff = SSTab1.Height - lngDiff
    If SSTab1.Tab = 0 Then
        CC.Height = Me.Height - (CC.top + 2000)
        CC.Width = Me.Width - (CC.left + 1705)
    Else
'        arvNeg.Height = Me.Height - (arvNeg.top + 1200)
'        arvNeg.Width = Me.Width - (arvNeg.left + 700)
    End If
    SSTab1.Width = Me.Width - (SSTab1.left + 400)
    
'    lblCF.top = SSTab1.top + SSTab1.Height - 1200
'    lblCFQTY.top = SSTab1.top + SSTab1.Height - 1200
'    lblCFVal.top = SSTab1.top + SSTab1.Height - 1200
'    lblCFPRQTY.top = SSTab1.top + SSTab1.Height - 1200
'
'    lblCalc.top = SSTab1.top + SSTab1.Height - 900
'    lblCalcQty.top = SSTab1.top + SSTab1.Height - 900
'    lblCalcVal.top = SSTab1.top + SSTab1.Height - 900
'    lblCalcPrQty.top = SSTab1.top + SSTab1.Height - 900
'
'    lblDiscr.top = SSTab1.top + SSTab1.Height - 600
'    lblDiscrQty.top = SSTab1.top + SSTab1.Height - 600
'    lblDiscrVal.top = SSTab1.top + SSTab1.Height - 600
'    lblDiscrPrQty.top = SSTab1.top + SSTab1.Height - 600
'    lblWarning.top = SSTab1.top + SSTab1.Height - 600

End Sub

Private Sub GridToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim DDLevel As Integer
 Dim Checked As Boolean
  
 Checked = (Button.Value = tbrPressed)
        CC.TitleSettings.Text = "TEST"

 With CC
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
 With CC
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
            CC.Descending = False
            CC.SortGridByFact SortAxis, Col, Row
   Case "desc"
       SortAxis = xda_vertical
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            CC.Descending = True
            CC.SortGridByFact SortAxis, Col, Row
   Case "Nosort"
       SortAxis = xda_vertical
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            CC.CancelFactSorting (SortAxis)
   Case "hasc"
       SortAxis = xda_horizontal
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            CC.Descending = False
            CC.SortGridByFact SortAxis, Col, Row
   Case "hdesc"
       SortAxis = xda_horizontal
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            CC.Descending = True
            CC.SortGridByFact SortAxis, Col, Row
   Case "hNosort"
       SortAxis = xda_horizontal
       Col = .CurrentCell.Col: Row = .CurrentCell.Row
            CC.CancelFactSorting (SortAxis)
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
 CC.ReportToFile FileName, "", FileFormat
 OpenDocument (FileName)
End Sub
Private Sub OpenDocument(f_name As String)
 Dim Scr_hDC As Long
 Scr_hDC = GetDesktopWindow()
 ShellExecute Scr_hDC, "Open", f_name, "", "", 1
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

Public Sub SaveContourCubeLayout(ltFile As String)
'Saving layout procedure
  Dim rsFields, Axis, Object, bInvertFilterSelection, Value, i, j, viewTotalsState, _
      viewGTotalsState, strExpand, fs
  rsFields = Array("Object", "Name", "Property", "Value")
  'Create an ADO recordset with 4 fields:
  Dim rs As New adodb.Recordset
  rs.Fields.Append rsFields(0), adBSTR, 10
  rs.Fields.Append rsFields(1), adBSTR, 50
  rs.Fields.Append rsFields(2), adVariant, 50
  rs.Fields.Append rsFields(3), adVariant, 255
  rs.Open
  rs.AddNew rsFields, Array("Cube", CC.Name, "RootAxis", CC.Cube.RootAxis)
  With CC
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
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .HAxis.GetSection(i).GetValue(j)
        Next j
        rs.AddNew rsFields, Array("Axis", "Horizontal", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    If .VAxis.Length > 0 Then
      For i = 0 To .VAxis.Length - 1
        strExpand = ""
        For j = 0 To .VAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .VAxis.GetSection(i).GetValue(j)
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
  Dim rs As New adodb.Recordset
  'First open the saved XML layout file
  rs.Open ltFile
  With CC
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
  If LCase(strAxis) = "horizontal" Then Set Axis = CC.HAxis Else Set Axis = CC.VAxis
  On Error Resume Next
  i = 0
  Do While i < Axis.Length
    j = 0
    Do While j <= UBound(aExpand, 1)
      If CStr(Axis.GetSection(i).GetValue(j)) <> aExpand(j) Then Exit Do
      j = j + 1
    Loop
    If j > UBound(aExpand, 1) Then Exit Do
    i = i + 1
  Loop
  If i < Axis.Length Then Axis.GetSection(i).Collapse UBound(aExpand, 1), True
  On Error GoTo 0
End Sub


