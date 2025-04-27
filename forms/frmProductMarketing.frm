VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form frmProductMarketing 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product marketing"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C4BCA4&
      Caption         =   "New rule"
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
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2100
      Width           =   1155
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Delete rule"
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
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2100
      Width           =   1155
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Rule"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2715
      Left            =   150
      TabIndex        =   1
      Top             =   2610
      Width           =   9180
      Begin VB.TextBox txtMinValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2430
         TabIndex        =   18
         Text            =   "0"
         Top             =   1680
         Width           =   1560
      End
      Begin VB.ComboBox cboCustomerGroup 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Text            =   "cboCustomerGroup"
         Top             =   600
         Width           =   2490
      End
      Begin VB.CheckBox chkActive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   7980
         TabIndex        =   7
         Top             =   1050
         Width           =   915
      End
      Begin VB.CheckBox chkNoDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "No discount allowable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   6585
         TabIndex        =   9
         Top             =   1635
         Width           =   2310
      End
      Begin VB.CheckBox chkID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Requires customer identification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   5715
         TabIndex        =   8
         Top             =   1335
         Width           =   3180
      End
      Begin VB.TextBox txtDescription 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3300
         MaxLength       =   95
         TabIndex        =   10
         Top             =   2190
         Width           =   4065
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7620
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2070
         Width           =   1500
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   330
         TabIndex        =   6
         Text            =   "0"
         Top             =   1680
         Width           =   900
      End
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmProductMarketing.frx":0000
         Left            =   3360
         List            =   "frmProductMarketing.frx":0002
         TabIndex        =   2
         Text            =   "cboProductType"
         Top             =   600
         Width           =   2490
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6420
         TabIndex        =   4
         Text            =   "cboSection"
         Top             =   600
         Width           =   2490
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Min. pre-disc sale value (Whole currency values only e.g. 800=R800)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   1665
         TabIndex        =   19
         Top             =   1200
         Width           =   3315
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   270
         TabIndex        =   17
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Description (only 20 characters will be shown on slip)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   525
         Left            =   300
         TabIndex        =   14
         Top             =   2100
         Width           =   2955
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   345
         TabIndex        =   13
         Top             =   1455
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Product type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3315
         TabIndex        =   12
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   6465
         TabIndex        =   3
         Top             =   345
         Width           =   1080
      End
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Bindings        =   "frmProductMarketing.frx":0004
      Height          =   1785
      Left            =   180
      OleObjectBlob   =   "frmProductMarketing.frx":0019
      TabIndex        =   0
      Top             =   225
      Width           =   9150
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   90
      Top             =   5310
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483635
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
      Caption         =   "Adodc1"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select and then double-click to edit"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   165
      TabIndex        =   16
      Top             =   2055
      Width           =   2955
   End
End
Attribute VB_Name = "frmProductMarketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private flgLoading As Boolean
Dim lngID As Long
Dim tlCustomerTypesActive As New z_TextList

Private Sub cboProductType_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    'oProd.SetProductTypeID oPC.Configuration.ProductTypes.Key(cboProductType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cboProductType_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cboCatHead_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    'oProd.SetCatalogueheadingID oPC.Configuration.CatalogueHeadings.Key(cboCatHead)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cboCatHead_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cmdClearAll_Click()
'If MsgBox("Confirm you want to erase all discount and
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("You wish to delete the marketing rule called " & G1.Columns(2), vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
Dim oSQL As New z_SQL
    If IsNull(G1.Bookmark) Then Exit Sub
    oSQL.RunSQL "DELETE FROM tMARKETING WHERE M_ID = " & FNN(G1.Columns(9))
    Adodc1.Refresh
End Sub

Private Sub cmdNew_Click()
    ClearTextFields
    fr1.Enabled = True
    cmdNew.Enabled = False
End Sub

Private Sub cmdSave_Click()
Dim PTID As Long
Dim SECTID As Long
Dim CUSTTYPEID As Long
Dim oSQL As New z_SQL
Dim lngResult As Long

    If cboProductType = "<ALL>" Then
        PTID = 0
    Else
        PTID = oPC.Configuration.ProductTypes.Key(cboProductType)
    End If
    If cboSection = "<ALL>" Then
        SECTID = 0
    Else
        SECTID = oPC.Configuration.Sections.Key(cboSection)
    End If
    If cboCustomerGroup = "<ALL>" Then
        CUSTTYPEID = 0
    Else
        CUSTTYPEID = tlCustomerTypesActive.Key(cboCustomerGroup)
    End If
    lngResult = oSQL.RunProc("sp_MarketingRules", Array(lngID, PTID, SECTID, Trim(txtDescription), CUSTTYPEID, FNDBL(IIf(txtDiscount = "", 0, txtDiscount)), FNN(IIf(txtMinValue = "", 0, txtMinValue)) * oPC.Configuration.DefaultCurrency.Divisor, IIf(chkID, "1", "0"), IIf(chkNoDiscount, "1", "0"), IIf(chkActive, "1", "0")), "")
    Adodc1.Refresh
    ClearTextFields
    lngID = 0
    cmdNew.Enabled = True
    fr1.Enabled = False
End Sub

Private Sub Form_Initialize()
    Set tlCustomerTypesActive = New z_TextList
    tlCustomerTypesActive.Load ltCustomerTypeActive, , "<ALL>"

End Sub

Private Sub Form_Load()
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    flgLoading = True
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "Select M_ID,M_PT_ID,M_SECTION_ID,M_DESCRIPTION,M_DISCOUNT,M_NoDiscountAllowable,M_IdentifyCustomer," _
        & " M_Active,M_CUSTTYPE_ID,M_MinValue/" & oPC.Configuration.DefaultCurrency.Divisor _
        & " as M_MinValue,d1.DICT_DESC AS SECTIONCODE,d2.DICT_DESC AS CUSTCODE,PT_CODE " _
        & " FROM  tMarketing LEFT JOIN tDICT d1 ON M_SECTION_ID = d1.DICT_ID " _
        & " LEFT JOIN tDICT d2 ON M_CUSTTYPE_ID = d2.DICT_ID  LEFT JOIN tPT ON M_PT_ID = PT_ID ORDER BY CustCode,PT_CODE,SECTIONCODE"
    Adodc1.ConnectionString = oPC.ConnectionString
    G1.DataSource = Me.Adodc1
    
    LoadCombo cboSection, oPC.Configuration.Sections
    cboSection = "<ALL>"
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    cboProductType = "<ALL>"
    LoadCombo cboCustomerGroup, tlCustomerTypesActive
    cboCustomerGroup = "<ALL>"
    fr1.Enabled = False
    
    flgLoading = False
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub


Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Cancel = True
End Sub

Private Sub G1_DblClick()
'Dim lngID As Long
Dim blnEdit As Boolean
    fr1.Enabled = True
    If IsNull(G1.Bookmark) Then Exit Sub
    lngID = G1.Columns(9)
    If G1.Columns(0) > "" Then
        Me.cboCustomerGroup = G1.Columns(0)
    Else
        Me.cboCustomerGroup = "<ALL>"
    End If
    If G1.Columns(1) > "" Then
        Me.cboProductType = oPC.Configuration.ProductTypes.Item(G1.Columns(10))
    Else
        Me.cboProductType = "<ALL>"
    End If
    If G1.Columns(2) > "" Then
        Me.cboSection = G1.Columns(2)
    Else
        Me.cboSection = "<ALL>"
    End If
    Me.txtDescription = G1.Columns(3)
    Me.txtDiscount = G1.Columns(4)
    Me.txtMinValue = G1.Columns(5)
    Me.chkID = IIf(G1.Columns(6) <> "0", 1, 0)
    Me.chkNoDiscount = IIf(G1.Columns(7) <> "0", 1, 0)
    Me.chkActive = IIf(G1.Columns(8) <> "0", 1, 0)
End Sub

Private Sub ClearTextFields()
    Me.cboProductType = "<ALL>"
    Me.cboSection = "<ALL>"
    Me.cboCustomerGroup = "<ALL>"
    Me.txtDescription = ""
    Me.txtDiscount = 0
    Me.txtMinValue = 0
    Me.chkID = 0
    Me.chkNoDiscount = 0
End Sub


Private Sub txtDiscount_Validate(Cancel As Boolean)
    Cancel = (IsNumeric(txtDiscount) = False)
End Sub
Private Sub txtMinValue_Validate(Cancel As Boolean)
    Cancel = (IsNumeric(txtMinValue) = False)
End Sub
