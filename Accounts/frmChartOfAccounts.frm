VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{3294EE21-3C6C-11D0-BADF-00201802BB87}#1.0#0"; "gtTree32.ocx"
Begin VB.Form frmChartOfAccounts 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Chart of accounts"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   Icon            =   "frmChartOfAccounts.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7455
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin DataTree.GTTree gtrCOA 
      Height          =   4050
      Left            =   8895
      TabIndex        =   1
      Top             =   1785
      Width           =   2805
      _Version        =   65536
      DefColCaptionBorderStyle=   3
      BeginProperty DefColCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DefFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ScrollTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ExtendTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PathSeparator   =   "\"
      LImgs.Count     =   1
      _ExtentX        =   4948
      _ExtentY        =   7144
      _StockProps     =   65
      BackColor       =   -2147483643
   End
   Begin DataTree.GTComboTree gtCOA 
      Height          =   315
      Left            =   8970
      TabIndex        =   0
      Top             =   495
      Width           =   2640
      _Version        =   65536
      DefColCaptionBorderStyle=   3
      BeginProperty DefColCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DefFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ScrollTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ExtendTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PathSeparator   =   "\"
      LImgs.Count     =   1
      _ExtentX        =   4657
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   2535
      Left            =   225
      OleObjectBlob   =   "frmChartOfAccounts.frx":038A
      TabIndex        =   2
      Top             =   480
      Width           =   5670
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   510
      Left            =   180
      Top             =   3150
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   900
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChartOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim rsAC As ADODB.Recordset




Private Sub Form_Load()
    On Error GoTo errHandler
 Dim i As Integer
 Dim vi As ValueItem
 
    oPC.OpenDBSHort
    Set rsAC = New ADODB.Recordset
    rsAC.CursorLocation = adUseClient
    rsAC.Open "SELECT ACC_Code,ACC_Description FROM tAccountCategory ORDER BY ACC_Description", oPC.COShort, adOpenStatic
    i = 0
    Do While Not rsAC.EOF
        i = i + 1
        Set vi = New ValueItem
        vi.Value = CLng(rsAC.Fields("ACC_Code"))
        vi.DisplayValue = FNS(rsAC.Fields("ACC_Description"))
        Grid.Columns(1).ValueItems.Add vi
        rsAC.MoveNext
    Loop
    Grid.Columns(1).ValueItems.Translate = True
    Grid.Columns(1).ValueItems.Presentation = dbgComboBox
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tAccount Order BY AC_Description", oPC.COShort, adOpenDynamic, adLockOptimistic
    Set Adodc1.Recordset = rs
    Set Grid.DataSource = Me.Adodc1
    Grid.Refresh
    Grid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmChartOfAccounts.Form_Load", , EA_NORERAISE
    HandleError
End Sub

