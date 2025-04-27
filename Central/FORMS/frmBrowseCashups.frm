VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmBrowseCashups 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Cashups"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   21300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   21300
   Begin TabDlg.SSTab SSTab1 
      Height          =   7635
      Left            =   180
      TabIndex        =   14
      Top             =   1245
      Width           =   21000
      _ExtentX        =   37042
      _ExtentY        =   13467
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Pivot view"
      TabPicture(0)   =   "frmBrowseCashups.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ContourCubeX"
      Tab(0).Control(1)=   "GridToolBar"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "List view"
      TabPicture(1)   =   "frmBrowseCashups.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GN"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdExporttoDelimited"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin CCubeX2.ContourCubeX ContourCubeX 
         Height          =   6000
         Left            =   -74835
         TabIndex        =   15
         Top             =   1260
         Width           =   20640
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
         CCubeXMetadata  =   $"frmBrowseCashups.frx":0038
      End
      Begin VB.CommandButton cmdExporttoDelimited 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Export"
         Height          =   345
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   420
         Width           =   1605
      End
      Begin TrueOleDBGrid60.TDBGrid GN 
         Height          =   6255
         Left            =   165
         OleObjectBlob   =   "frmBrowseCashups.frx":04E0
         TabIndex        =   16
         Top             =   825
         Width           =   20595
      End
      Begin MSComctlLib.Toolbar GridToolBar 
         Height          =   660
         Left            =   -74850
         TabIndex        =   17
         Top             =   465
         Width           =   7305
         _ExtentX        =   12885
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Filter"
      Height          =   1005
      Left            =   180
      TabIndex        =   0
      Top             =   105
      Width           =   10695
      Begin VB.TextBox txtTillpoint 
         Height          =   330
         Left            =   1440
         TabIndex        =   12
         Top             =   450
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dtpSince 
         Height          =   330
         Left            =   6510
         TabIndex        =   9
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56950785
         CurrentDate     =   40174
      End
      Begin VB.CommandButton cmdLoadFromFilter 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Filter"
         Height          =   525
         Left            =   9495
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   1005
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmBrowseCashups.frx":ECF7
         Left            =   4575
         List            =   "frmBrowseCashups.frx":ED0A
         TabIndex        =   6
         Top             =   465
         Width           =   1440
      End
      Begin VB.TextBox txtDiscrepancy 
         Height          =   330
         Left            =   3060
         TabIndex        =   3
         Top             =   450
         Width           =   1440
      End
      Begin VB.ComboBox cboStores 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   465
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpUntil 
         Height          =   330
         Left            =   7950
         TabIndex        =   11
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56950785
         CurrentDate     =   40174
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Store"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   255
         TabIndex        =   13
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "until"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   8205
         TabIndex        =   10
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblSince 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "from"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6765
         TabIndex        =   8
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4845
         TabIndex        =   5
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discrepancy > than"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3090
         TabIndex        =   4
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label lblStores 
         BackStyle       =   0  'Transparent
         Caption         =   "Tillpoint"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1665
         TabIndex        =   2
         Top             =   225
         Width           =   795
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   21960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12825
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   11445
      Top             =   225
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
            Picture         =   "frmBrowseCashups.frx":ED3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":F3CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":FA5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":100F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":10783
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":10E15
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":114A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":11659
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":1180B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":119BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":11B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":11D15
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":123A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":12A39
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":130CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":1375D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":13AF7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList HotImageList 
      Left            =   13980
      Top             =   180
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
            Picture         =   "frmBrowseCashups.frx":13E91
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":14043
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":141F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":143A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":14559
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":1470B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":148BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":14A6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":14C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":14DD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":14F85
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":1512B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":152DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":1548F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCashups.frx":15641
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowseCashups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCashups As New c_Cashups
Dim XA As New XArrayDB
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

Private Sub cmdExporttoDelimited_Click()
    On Error GoTo errHandler
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim strExecutable As String

    Screen.MousePointer = vbHourglass
    
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\TEMP"
    End If
    sFile = oPC.SharedFolderRoot & "\TEMP\Cashups.csv"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    Me.GN.ExportToDelimitedFile sFile, , ","
    
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
           ShellandWait """" & strExecutable & """" & " " & sFile
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 70 Then
        MsgBox "It look like you already have an exported document open. Close it before re-exporting.", vbInformation + vbOKOnly, "Can't do this"
        Err.Clear
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCashups.cmdExporttoDelimited_Click"
    HandleError
End Sub

Private Sub cmdLoadFromFilter_Click()
    LoadCashupsFromDB
    LoadGrid
    LoadPT
End Sub

Private Sub Form_Load()
    Me.top = 500
    Me.left = 300
    Me.Width = 11100
    Me.Height = 6000
    InitializeFilterCommands
    SetGridLayout Me.GN, Me.Name
    SetFormSize Me
    
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

Private Sub InitializeFilterCommands()

Dim oStore As a_Store

    With cboStores
        .Clear
        .AddItem ""
        
        For Each oStore In oPC.Configuration.Stores
            .AddItem oStore.code
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    Me.dtpUntil = DateAdd("yyyy", 1, Date)
End Sub
Private Sub LoadCashupsFromDB()
    Set rs = New ADODB.Recordset
    oCashups.LoadRecordset Me.cboStores, txtDiscrepancy, Me.txtTillpoint, Me.dtpSince, dtpUntil, Me.cboStatus, rs
    oCashups.Load Me.cboStores, txtDiscrepancy, Me.txtTillpoint, Me.dtpSince, dtpUntil, Me.cboStatus
End Sub
Private Sub LoadGrid()
Dim objItem As d_Cashup
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, oCashups.Count, 1, 33
    For i = 1 To oCashups.Count
        With objItem
            XA.Value(i, 1) = oCashups.Item(i).branchcode
            XA.Value(i, 2) = oCashups.Item(i).Tillpoint
            XA.Value(i, 3) = oCashups.Item(i).OpenSessionTimeF
            XA.Value(i, 4) = oCashups.Item(i).Status
            XA.Value(i, 5) = oCashups.Item(i).StatusDateF
            XA.Value(i, 6) = oCashups.Item(i).StatusSignature
            XA.Value(i, 7) = oCashups.Item(i).OpeningFloatF
            XA.Value(i, 8) = oCashups.Item(i).ClosingFloatF
            XA.Value(i, 9) = oCashups.Item(i).CashF
            XA.Value(i, 10) = oCashups.Item(i).ChequesF
            XA.Value(i, 11) = oCashups.Item(i).CreditCardsF
            XA.Value(i, 12) = oCashups.Item(i).DebitCardsF
            XA.Value(i, 13) = oCashups.Item(i).DirectDepositsF
            XA.Value(i, 14) = oCashups.Item(i).VouchersRedeemedF
            XA.Value(i, 15) = oCashups.Item(i).DiscrepancyAllF
            XA.Value(i, 16) = oCashups.Item(i).Explanation
            XA.Value(i, 17) = oCashups.Item(i).WagesF
            XA.Value(i, 18) = oCashups.Item(i).LeavePayF
            XA.Value(i, 19) = oCashups.Item(i).SickLeaveF
            XA.Value(i, 20) = oCashups.Item(i).TotalSalesF
            XA.Value(i, 21) = oCashups.Item(i).COGSF
            XA.Value(i, 22) = oCashups.Item(i).RetainedF
            XA.Value(i, 23) = oCashups.Item(i).ReturnedF
            XA.Value(i, 24) = oCashups.Item(i).GiftVouchersSoldF
            XA.Value(i, 25) = oCashups.Item(i).OtherVouchersSoldF
            XA.Value(i, 26) = oCashups.Item(i).TotalSales
            XA.Value(i, 27) = ""
            XA.Value(i, 28) = oCashups.Item(i).xid
            XA.Value(i, 29) = oCashups.Item(i).StatusDate
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
    GN.Array = XA
    GN.ReBind
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
Dim lngDiffH As Long
    Me.SSTab1.Width = NonNegative_Lng(Me.Width - 500)
    GN.Width = NonNegative_Lng(Me.SSTab1.Width - 500)
    Me.ContourCubeX.Width = NonNegative_Lng(Me.SSTab1.Width - 500)
    lngDiff = GN.Height
    SSTab1.Height = NonNegative_Lng(Me.Height - (Me.top + 1600))
    GN.Height = NonNegative_Lng(Me.Height - (GN.top + 2400))
    ContourCubeX.Height = NonNegative_Lng(Me.Height - (GN.top + 2400))

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Form_Resize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mnuSaveLayout
End Sub

Private Sub txtDiscrepancy_Validate(Cancel As Boolean)
    Cancel = (Not IsNumeric(txtDiscrepancy)) And txtDiscrepancy <> ""
End Sub

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.GN, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.mnuSaveLayout"
End Sub

'==============================================================
'==============================================================
'==============================================================
'==============================================================
'==============================================================
Private Sub LoadPT()
    On Error GoTo errHandler
    Dim iView As IContourView
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
'        .Dims.Add("Rank", "Rank_Rank", , xda_vertical).MoveTo xda_vertical, 0
        .Dims.Add("yr", "Yr", , xda_vertical).MoveTo xda_outside
        .Dims.Add("mth", "Mth", , xda_vertical).MoveTo xda_outside
        .Dims.Add("wk", "Wk", , xda_vertical).MoveTo xda_outside
        .Dims.Add("Till", "CU_Tillpoint", , xda_vertical).MoveTo xda_outside
        
        .Dims.Add("Br Code", "CU_BranchCode", , xda_vertical).MoveTo xda_vertical, 0
        .Dims.Add("StartSession", "CU_OpenSessionTime", , xda_vertical).MoveTo xda_vertical, 1
        .Dims.Add("Status", "STATUS", , xda_vertical).MoveTo xda_vertical, 2
        .Dims.Add("StatusDate", "StatusDate", , xda_vertical).MoveTo xda_vertical, 3
        .Dims.Add("Signature", "StatusSignature", , xda_vertical).MoveTo xda_vertical, 4
        
'        .Dims.Add("COMBO", "Title", , xda_vertical).MoveTo xda_outside
'        .Dims.Add("Cost", "P_Cost", , xda_vertical).MoveTo xda_outside
'        .Dims.Add("SP", "P_SP", , xda_vertical).MoveTo xda_outside
'        .Dims.Add("AcnoCustomer", "Acno", , xda_vertical).MoveTo xda_outside, 1
'        .Dims.Add("Acno", "CustomerAcno", , xda_vertical).MoveTo xda_outside, 1
'        .Dims.Add("EXCHANGENUMBER", "EXCHANGENUMBER", , xda_vertical).MoveTo xda_outside, 1
'        .Dims.Add("HBr", "STORE_CODE", , xda_vertical).MoveTo xda_outside, 1
        
'        .BaseFacts.Add "bfTotalSales", "CU_TotalSales"
'        .Facts.Add "TotalSales", "bfTotalSales", xfaa_SUM
'        ContourCubeX.Facts(0).Appearance.Format = "###,##0.00"
'
'        .BaseFacts.Add "bfCOGS", "CU_COGS"
'        .Facts.Add "COGS", "bfCOGS", xfaa_SUM
'        ContourCubeX.Facts(1).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfCash", "CU_Cash"
        .Facts.Add "Cash", "bfCash", xfaa_SUM
        ContourCubeX.Facts(0).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfCheques", "CU_Cheques"
        .Facts.Add "Cheques", "bfCheques", xfaa_SUM
        ContourCubeX.Facts(1).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfCreditCards", "CU_CreditCards"
        .Facts.Add "Cr Cards", "bfCreditCards", xfaa_SUM
        ContourCubeX.Facts(2).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfDebitCards", "CU_DebitCards"
        .Facts.Add "Db Cards", "bfDebitCards", xfaa_SUM
        ContourCubeX.Facts(3).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfDirectDeposits", "CU_DirectDeposits"
        .Facts.Add "Dir dep", "bfDirectDeposits", xfaa_SUM
        ContourCubeX.Facts(4).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfVouchersRedeemed", "CU_Vouchers"
        .Facts.Add "V red", "bfVouchersRedeemed", xfaa_SUM
        ContourCubeX.Facts(5).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfTotalDiscrepancy", "TotalDiscrepancy"
        .Facts.Add "Tot discr", "bfTotalDiscrepancy", xfaa_SUM
        ContourCubeX.Facts(6).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfWages", "CU_Wages"
        .Facts.Add "Wages", "bfWages", xfaa_SUM
        ContourCubeX.Facts(7).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfLeavePay", "CU_LeavePay"
        .Facts.Add "Leave pay", "bfLeavePay", xfaa_SUM
        ContourCubeX.Facts(8).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfSickLeave", "CU_SickLeave"
        .Facts.Add "Sick leave", "bfSickLeave", xfaa_SUM
        ContourCubeX.Facts(1).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfTotalSales", "CU_TotalSales"
        .Facts.Add "Tot Sal", "bfTotalSales", xfaa_SUM
        ContourCubeX.Facts(9).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfCOGS", "CU_COGS"
        .Facts.Add "COGS", "bfCOGS", xfaa_SUM
        ContourCubeX.Facts(10).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfRetained", "CU_Retained"
        .Facts.Add "Retained", "bfRetained", xfaa_SUM
        ContourCubeX.Facts(11).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfReturned", "CU_Returned"
        .Facts.Add "Returned", "bfReturned", xfaa_SUM
        ContourCubeX.Facts(12).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfGiftVouchersSold", "CU_GiftVouchersSold"
        .Facts.Add "Gift V sold", "bfGiftVouchersSold", xfaa_SUM
        ContourCubeX.Facts(13).Appearance.Format = "###,##0.00"
        
        .BaseFacts.Add "bfOtherVouchersSold", "CU_OtherVouchersSold"
        .Facts.Add "Oth V issd", "bfOtherVouchersSold", xfaa_SUM
        ContourCubeX.Facts(14).Appearance.Format = "###,##0.00"
        
        For Each Fact In ContourCubeX.Facts
          Fact.Visible = True
        Next
        Set rs.ActiveConnection = Nothing
        '.DataSourceType = xcdt_Recordset
        .Open rs
        ContourCubeX.Dims(4).Descending = True
        ContourCubeX.Dims(4).NoTotals = True
        ContourCubeX.Dims(5).NoTotals = True
        ContourCubeX.Dims(6).NoTotals = True
        ContourCubeX.Dims(7).NoTotals = True
        ContourCubeX.Dims(8).NoTotals = True
'        iView.Dims(6).NoTotals = True
    End With
  
    AfterOpen
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCashups.LoadPT"
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

