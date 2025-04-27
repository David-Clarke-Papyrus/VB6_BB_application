VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgedAppros 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Report filter"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   7455
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Appros issued"
      ForeColor       =   &H8000000D&
      Height          =   1890
      Left            =   90
      TabIndex        =   7
      Top             =   180
      Width           =   4995
      Begin VB.CheckBox chkApproAll 
         BackColor       =   &H00D3D3CB&
         Caption         =   "All Customers       or"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   435
         TabIndex        =   13
         Top             =   990
         Width           =   2130
      End
      Begin VB.CommandButton cmdSelectCustomer 
         BackColor       =   &H00C4BCA4&
         Cancel          =   -1  'True
         Caption         =   "&Select customer"
         Height          =   465
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   900
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1110
         TabIndex        =   8
         Top             =   345
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
         Format          =   58130433
         CurrentDate     =   37421
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   3285
         TabIndex        =   9
         Top             =   345
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
         Format          =   58130433
         CurrentDate     =   37421
      End
      Begin VB.Label lblCustomer 
         BackStyle       =   0  'Transparent
         Height          =   330
         Left            =   135
         TabIndex        =   14
         Top             =   1395
         Width           =   4665
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2640
         TabIndex        =   11
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "between"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   135
         TabIndex        =   10
         Top             =   390
         Width           =   840
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
      Left            =   6330
      Picture         =   "frmAgedAppros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1425
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
      Left            =   5310
      Picture         =   "frmAgedAppros.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1425
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H8000000D&
      Height          =   1005
      Left            =   5160
      TabIndex        =   2
      Top             =   180
      Width           =   2175
      Begin VB.OptionButton optAtCost 
         BackColor       =   &H00D3D3CB&
         Caption         =   "At Cost"
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
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optAtList 
         BackColor       =   &H00D3D3CB&
         Caption         =   "At sell. price"
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
         Left            =   225
         TabIndex        =   3
         Top             =   210
         Width           =   1785
      End
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   390
      Left            =   2025
      OleObjectBlob   =   "frmAgedAppros.frx":0714
      TabIndex        =   0
      Top             =   4545
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
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
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   4605
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "frmAgedAppros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim strCustomerName As String
Dim lngPTID As Long
Dim strPT As String
Dim bCancelled As Boolean

Public Sub Component(pCaption As String, Optional ShowCostListOption As Boolean)
    Me.Caption = pCaption
    Me.Frame1.Visible = ShowCostListOption
End Sub

Private Sub SetupPT()
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add "Seesafe"
    cboProductType.Columns(0).Width = 190
    cboProductType.Columns(1).Width = 0
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate
End Sub


'Private Sub cmdAll_Click()
'    strCustomerName = "<ALL>"
'    lngTPID = 0
'    txtCustomer = strCustomerName
'End Sub

Private Sub chkApproAll_Click()
    If chkApproAll = 1 Then
        Me.lblCustomer.Caption = ""
        lngTPID = 0
    End If

End Sub

Private Sub cmdClose_Click()
    bCancelled = True
    Me.Hide
End Sub



Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngTPID As Long
'Dim frm As frmCustomersTA
Dim strSQL As String
Dim rs As adodb.Recordset
Dim frmR As frmApprosPT
Dim dte1 As Date
Dim dte2 As Date
Dim bAtCost As Boolean

    Me.MousePointer = vbHourglass
    If Me.Cancelled Then
        DoEvents
        Exit Sub
    End If
    If chkApproAll.Value = 0 And lngTPID = 0 Then
        MsgBox "Enter a customer or check All Customers.", vbOKOnly + vbInformation, _
                        "Papyrus Reports - Status"
        GoTo EXIT_Handler
    End If
    Screen.MousePointer = vbHourglass
    dte1 = Me.StartDate
    dte2 = Me.EndDate
    lngTPID = Me.CustomerID
    bAtCost = Me.AtCost
    Unload Me
    If bAtCost Then
        If lngTPID > 0 Then
            strSQL = "SELECT * FROM vOSAppros_AtCost WHERE   TP_ID = " & lngTPID
        Else
            'strSQL = "SELECT * FROM vOSAppros_AtCost WHERE DocDate BETWEEN '" & ReverseDate(dte1) & "' AND '" & ReverseDate(dte2) & "'"
            strSQL = "SELECT * FROM ReportAppro WHERE TR_DATE BETWEEN dbo.StartOfDay('" & ReverseDate(dte1) & "') AND dbo.EndOfDay('" & ReverseDate(dte2) & "') AND ExtRetailOSIncVAT > 0"
        End If
    Else
        If lngTPID > 0 Then
            strSQL = "SELECT * FROM vOSAppros WHERE   TP_ID = " & lngTPID
        Else
            strSQL = "SELECT * FROM vOSAppros WHERE DocDate BETWEEN '" & ReverseDate(dte1) & "' AND '" & ReverseDate(dte2) & "'"
        End If
    End If
    Set rs = New adodb.Recordset
    Forms(0).SB1.Panels(1).Text = "Loading . . . "
    DoEvents
    Forms(0).SB1.Panels(1).Text = ""
    rs.Open strSQL, oPC.CO
    Set frmR = New frmApprosPT
    frmR.Component rs
    Set rs = Nothing
    Unload Me
    frmR.Show
    

EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAgedAppros.cmdOK_Click"
End Sub

'Private Sub cmdCust_Click()
'Dim frm As frmBrowseCustomers2
'    Set frm = New frmBrowseCustomers2
'    frm.Show vbModal
'    lngTPID = frm.CustomerID
'    strCustomerName = frm.CustomerName
'    txtCustomer = strCustomerName
'    Unload frm
'    If lngTPID = 0 Then Exit Sub
'
'End Sub

Private Sub cmdSelectCustomer_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    strCustomerName = left(frm.CustomerName, 40) & IIf(frm.Accnum > "", " (" & frm.Accnum & ")", "")
    Me.lblCustomer.Caption = strCustomerName
    Unload frm
    If lngTPID = 0 Then Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmaGEDaPPROS.cmdSelectCustomer_Click"
End Sub

Private Sub Form_Initialize()
Dim ar() As String
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate

End Sub

Private Sub Form_Load()
    SetupPT
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Width = 7700
    Height = 2700
    left = 500
    top = 1000
    bCancelled = False
End Sub


Private Sub cboProductType_SelectionChanged()
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    strPT = cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0)
End Sub

Property Get CustomerID() As Long
    CustomerID = lngTPID
End Property
Property Get PTID() As Long
    PTID = lngPTID
End Property
Property Get StartDate() As Date
    StartDate = CDate(dtpFrom.Value)
End Property
Property Get EndDate() As Date
    EndDate = CDate(dtpTo.Value)
End Property
Property Get CustomerName() As String
    CustomerName = strCustomerName
End Property
Property Get PTName() As String
    PTName = strPT
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Public Property Get AtCost() As Boolean
    AtCost = optAtCost
End Property

