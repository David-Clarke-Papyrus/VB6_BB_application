VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSuppliersTA 
   BackColor       =   &H00D3D3CB&
   Caption         =   "SuppliersTA"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   5190
   Begin VB.CheckBox chkLDP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Use last delivered cost (not weighted average)"
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   450
      TabIndex        =   12
      Top             =   1935
      Width           =   2415
   End
   Begin VB.CheckBox chkExVAT 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Values Ex V.A.T."
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   450
      TabIndex        =   11
      Top             =   1605
      Width           =   1635
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
      Left            =   2910
      Picture         =   "frmSuppliersTAs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1635
      Width           =   1000
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
      Left            =   3930
      Picture         =   "frmSuppliersTAs.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1635
      Width           =   1000
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   660
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   390
      Left            =   1695
      TabIndex        =   4
      Top             =   960
      Width           =   2550
   End
   Begin VB.CommandButton cmdSupp 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select &supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   945
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   225
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
      Left            =   3555
      TabIndex        =   1
      Top             =   225
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
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   390
      Left            =   1665
      OleObjectBlob   =   "frmSuppliersTAs.frx":0714
      TabIndex        =   5
      Top             =   3045
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "between"
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
      Left            =   405
      TabIndex        =   8
      Top             =   255
      Width           =   840
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
      Left            =   120
      TabIndex        =   6
      Top             =   3105
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Left            =   2910
      TabIndex        =   2
      Top             =   255
      Width           =   435
   End
End
Attribute VB_Name = "frmSuppliersTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim strSupplierName As String
Dim lngPTID As Long
Dim strPT As String
Dim bCancelled As Boolean

Public Sub Component(pCaption As String)
    Me.Caption = pCaption
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


Private Sub cmdAll_Click()
    strSupplierName = "<ALL>"
    lngTPID = 0
    txtSupplier = strSupplierName
End Sub

Private Sub cmdClose_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub


Private Sub cmdSupp_Click()
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    strSupplierName = frm.SupplierName
    txtSupplier = strSupplierName
    Unload frm
    If lngTPID = 0 Then Exit Sub

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
    Width = 5500
    Height = 3200
    left = 500
    top = 1000
    
    bCancelled = False
End Sub


Private Sub cboProductType_SelectionChanged()
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    strPT = cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0)
End Sub

Property Get SupplierID() As Long
    SupplierID = lngTPID
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
Property Get SupplierName() As String
    SupplierName = strSupplierName
End Property
Property Get PTName() As String
    PTName = strPT
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
