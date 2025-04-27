VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmBrowseCOLS 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales orders outstanding"
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
      TabIndex        =   8
      Top             =   1245
      Width           =   21000
      _ExtentX        =   37042
      _ExtentY        =   13467
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmBrowseCOLS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GridToolBar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ContourCubeX"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmBrowseCOLS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GN"
      Tab(1).Control(1)=   "cmdExporttoDelimited"
      Tab(1).ControlCount=   2
      Begin CCubeX2.ContourCubeX ContourCubeX 
         Height          =   6000
         Left            =   165
         TabIndex        =   9
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
         CCubeXMetadata  =   $"frmBrowseCOLS.frx":0038
      End
      Begin VB.CommandButton cmdExporttoDelimited 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Export"
         Height          =   345
         Left            =   -74835
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   420
         Width           =   1605
      End
      Begin TrueOleDBGrid60.TDBGrid GN 
         Height          =   6255
         Left            =   -74835
         OleObjectBlob   =   "frmBrowseCOLS.frx":04E0
         TabIndex        =   10
         Top             =   825
         Width           =   20595
      End
      Begin MSComctlLib.Toolbar GridToolBar 
         Height          =   660
         Left            =   150
         TabIndex        =   11
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
      Width           =   5820
      Begin MSComCtl2.DTPicker dtpSince 
         Height          =   330
         Left            =   1485
         TabIndex        =   4
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   40174
      End
      Begin VB.CommandButton cmdLoadFromFilter 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Filter"
         Height          =   525
         Left            =   4470
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   1005
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
         Left            =   2925
         TabIndex        =   6
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   40174
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Store"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "until"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3180
         TabIndex        =   5
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblSince 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "from"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1740
         TabIndex        =   3
         Top             =   225
         Width           =   600
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
            Picture         =   "frmBrowseCOLS.frx":96DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":9D6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":A3FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":AA91
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":B123
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":B7B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":BE47
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":BFF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":C1AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":C35D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":C50F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":C6B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":CD47
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":D3D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":DA6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":E0FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":E497
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
            Picture         =   "frmBrowseCOLS.frx":E831
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":E9E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":EB95
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":ED47
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":EEF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":F0AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":F25D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":F40F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":F5C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":F773
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":F925
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":FACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":FC7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":FE2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseCOLS.frx":FFE1
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowseCOLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCOLS As New c_COLS
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
            F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus, 10000
    End If
'errHandler:
'    ErrPreserve
'    If Err = 70 Then
'        MsgBox "It look like you already have an exported document open. Close it before re-exporting.", vbInformation + vbOKOnly, "Can't do this"
'        Err.Clear
'        Exit Sub
'    End If
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCashups.cmdExporttoDelimited_Click"
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.cmdExporttoDelimited_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLoadFromFilter_Click()
    On Error GoTo errHandler
    LoadCOLSFromDB
    LoadGrid
    LoadPT
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.cmdLoadFromFilter_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Me.top = 500
    Me.left = 300
    Me.Width = 11100
    Me.Height = 6000
    InitializeFilterCommands
    SetGridLayout Me.GN, Me.Name
    SetFormSize Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub ContourCubeX_OnCubeLoaded()
    On Error GoTo errHandler
    CheckVisible
    CheckEnabled
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesPT.ContourCubeX_OnCubeLoaded"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.ContourCubeX_OnCubeLoaded", , EA_NORERAISE
    HandleError
End Sub

Private Sub InitializeFilterCommands()
    On Error GoTo errHandler

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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.InitializeFilterCommands"
End Sub
Private Sub LoadCOLSFromDB()
    On Error GoTo errHandler
    Set rs = New ADODB.Recordset
    oCOLS.LoadRecordset Me.cboStores, Me.dtpSince, dtpUntil, rs
    oCOLS.Load Me.cboStores, Me.dtpSince, dtpUntil
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.LoadCOLSFromDB"
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItem As d_Cashup
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, oCOLS.Count, 1, 33
    For i = 1 To oCOLS.Count
        With objItem
            XA.Value(i, 1) = oCOLS.Item(i).branchcode
            XA.Value(i, 2) = oCOLS.Item(i).DocumentCode
            XA.Value(i, 3) = oCOLS.Item(i).DocumentCaptureDateF
            XA.Value(i, 4) = oCOLS.Item(i).DocumentIssueDateF
            XA.Value(i, 5) = oCOLS.Item(i).DOCUMENTCAPTUREDBY
            XA.Value(i, 6) = oCOLS.Item(i).CustomerName
            XA.Value(i, 7) = oCOLS.Item(i).CustomerAcno
            XA.Value(i, 8) = oCOLS.Item(i).CustomerPhone
            XA.Value(i, 9) = oCOLS.Item(i).ProductEAN
            XA.Value(i, 10) = oCOLS.Item(i).ProductTitle
            XA.Value(i, 11) = oCOLS.Item(i).SellingPriceF
            XA.Value(i, 12) = oCOLS.Item(i).OrderlineQtyF
            XA.Value(i, 13) = oCOLS.Item(i).OrderlineQtyDispatchedF
            XA.Value(i, 14) = oCOLS.Item(i).OrderlineQtyOutstandingF
            XA.Value(i, 15) = oCOLS.Item(i).OrderlineDiscountF
            XA.Value(i, 16) = oCOLS.Item(i).OrderlineRef
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
    GN.Array = XA
    GN.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.LoadGrid"
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

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCOLAllocation.Form_Resize", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.Form_Resize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.GN, Me.Name, Me.Height, Me.Width
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCOLAllocation.mnuSaveLayout"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.mnuSaveLayout"
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
        
        .Dims.Add("Br Code", "COL_STORECODE", , xda_vertical).MoveTo xda_vertical, 0
        .Dims.Add("Captured", "COL_DOCUMENTCAPTUREDATE", , xda_vertical).MoveTo xda_vertical, 1
        .Dims.Add("Document", "COL_DOCUMENTCODE", , xda_vertical).MoveTo xda_vertical, 2
        .Dims.Add("Issued", "COL_DOCUMENTISSUEDATE", , xda_vertical).MoveTo xda_vertical, 3
        .Dims.Add("By", "COL_DOCUMENTCAPTUREDBY", , xda_vertical).MoveTo xda_vertical, 4
        .Dims.Add("Customer", "COL_CUSTOMERNAME", , xda_vertical).MoveTo xda_vertical, 4
        .Dims.Add("A/c no", "COL_CUSTOMERACNO", , xda_vertical).MoveTo xda_vertical, 4
        
        ContourCubeX.Dims(4).Descending = True
        ContourCubeX.Dims(4).NoTotals = True
        ContourCubeX.Dims(5).NoTotals = True
        ContourCubeX.Dims(6).NoTotals = True
        ContourCubeX.Dims(7).NoTotals = True
        ContourCubeX.Dims(8).NoTotals = True
        ContourCubeX.Dims(9).NoTotals = True
      '  ContourCubeX.Dims(10).NoTotals = True
        
        .BaseFacts.Add "bfOLQty", "COL_OrderlineQty"
        .Facts.Add "Qty", "bfOLQty", xfaa_SUM
        ContourCubeX.Facts(0).Appearance.Format = "###,##0"
        
        .BaseFacts.Add "bfOLQtyDispatched", "COL_OrderlineQtyDispatched"
        .Facts.Add "Dispatched", "bfOLQtyDispatched", xfaa_SUM
        ContourCubeX.Facts(1).Appearance.Format = "###,##0"
        
        .BaseFacts.Add "bfCOLOutstanding", "COL_OrderlineQtyOutstanding"
        .Facts.Add "O/S", "bfCOLOutstanding", xfaa_SUM
        ContourCubeX.Facts(2).Appearance.Format = "###,##0"
        
        

        
        For Each Fact In ContourCubeX.Facts
          Fact.Visible = True
        Next
        Set rs.ActiveConnection = Nothing
        '.DataSourceType = xcdt_Recordset
        .Open rs
'        iView.Dims(6).NoTotals = True
    End With
  
    AfterOpen
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCashups.LoadPT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.LoadPT"
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
 ContourCubeX.Visible = ContourCubeX.Active
 CheckEnabled
 CheckVisible
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesPT.AfterOpen"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.AfterOpen"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesPT.CloseCube"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.CloseCube"
End Sub
Private Sub CheckEnabled()
    On Error GoTo errHandler
 Dim i As Integer
 'Check if controls are enabled or not
 For i = 1 To GridToolBar.Buttons.Count
  GridToolBar.Buttons(i).Enabled = ContourCubeX.Active
 Next i
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesPT.CheckEnabled"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.CheckEnabled"
End Sub

Private Sub CheckVisible()
    On Error GoTo errHandler
 ContourCubeX.Visible = True 'ContourCubeX.Active
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesPT.CheckVisible"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.CheckVisible"
End Sub
Private Sub GridToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.GridToolBar_ButtonClick(Button)", Button, EA_NORERAISE
    HandleError
End Sub

Private Sub GridToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.GridToolBar_ButtonMenuClick(ButtonMenu)", ButtonMenu, EA_NORERAISE
    HandleError
End Sub
Private Sub ExportCube(FileName As String, FileFormat As TxOlapReportType, FileType As String)
    On Error GoTo errHandler
 'Export OLAP-report to Excel, Word, HTML as file in html format
 FileName = FileName + "." + FileType
 ContourCubeX.ReportToFile FileName, "", FileFormat
 OpenDocument (FileName)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.ExportCube(FileName,FileFormat,FileType)", Array(FileName, FileFormat, _
         FileType)
End Sub
Private Sub OpenDocument(f_name As String)
    On Error GoTo errHandler
 Dim Scr_hDC As Long
 Scr_hDC = GetDesktopWindow()
 ShellExecute Scr_hDC, "Open", f_name, "", "", 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.OpenDocument(f_name)", f_name
End Sub


Public Sub SaveContourCubeLayout(ltFile As String)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.SaveContourCubeLayout(ltFile)", ltFile
End Sub

Sub LoadContourcubeLayout(ltFile As String)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.LoadContourcubeLayout(ltFile)", ltFile
End Sub

Sub ExpandSection(strAxis As String, strExpand As String)
    On Error GoTo errHandler
'This procedure restores saved state of an axis section
'It searches for given combination of dim values along the axis,
'and expands the section found
  Dim Axis As IViewAxis, i, j, aExpand
  aExpand = Split(strExpand, Chr(10))
  If LCase(strAxis) = "horizontal" Then Set Axis = ContourCubeX.HAxis Else Set Axis = ContourCubeX.VAxis
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
    ErrorIn "frmBrowseCOLS.ExpandSection(strAxis,strExpand)", Array(strAxis, strExpand)
End Sub
Private Sub LoadFormat()
    On Error GoTo errHandler
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Load Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
  CommonDialog1.CancelError = True
  CommonDialog1.ShowOpen
  If Err.Number = cdlCancel Then
    Exit Sub
  Else
    LoadContourcubeLayout CommonDialog1.FileName
  End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.LoadFormat"
End Sub

Private Sub SaveFormat()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\CubeFormats") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\CubeFormats")
    End If
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Save Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
  CommonDialog1.CancelError = True
  CommonDialog1.ShowSave
  If Err.Number = cdlCancel Then
    Exit Sub
  Else
    If Trim(CommonDialog1.FileName) <> "" Then SaveContourCubeLayout CStr(CommonDialog1.FileName)
  End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.SaveFormat"
End Sub
Private Sub GN_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
 '   If ColIndex = 2 Then ColIndex = 4
'    If ColIndex = 0 Then
'        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
'    ElseIf ColIndex = 2 Then
'        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 4, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
'    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 3, XORDER_DESCEND, XTYPE_DATE
   ' End If
    
    GN.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 11, 12, 13, 14, 15
            GetRowType = XTYPE_NUMBER
        Case 3, 4
            GetRowType = XTYPE_DATE
        Case Else
            GetRowType = XTYPE_STRING
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOLS.GetRowType(ColIndex)", ColIndex
End Function

