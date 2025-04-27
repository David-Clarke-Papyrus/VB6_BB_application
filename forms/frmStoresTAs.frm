VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStoresTA 
   BackColor       =   &H00D3D3CB&
   Caption         =   "SuppliersTA"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2850
      Picture         =   "frmStoresTAs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2490
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
      Left            =   3870
      Picture         =   "frmStoresTAs.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2490
      Width           =   1000
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
      Format          =   58785793
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3555
      TabIndex        =   1
      Top             =   210
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
      Format          =   58785793
      CurrentDate     =   37421
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   390
      Left            =   1290
      OleObjectBlob   =   "frmStoresTAs.frx":0714
      TabIndex        =   3
      Top             =   1620
      Width           =   3255
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboStore 
      Height          =   330
      Left            =   1320
      OleObjectBlob   =   "frmStoresTAs.frx":1ABE
      TabIndex        =   8
      Top             =   975
      Visible         =   0   'False
      Width           =   2895
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
      TabIndex        =   5
      Top             =   240
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
      Left            =   -255
      TabIndex        =   4
      Top             =   1680
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
Attribute VB_Name = "frmStoresTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim lngStoreID As Long
Dim strStoreName As String
Dim lngPTID As Long
Dim strPT As String
Dim bCancelled As Boolean
Dim tlStores As z_TextList

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
    
    
    cboStore.BeginUpdate
    cboStore.WidthList = 190
    cboStore.HeightList = 162
    cboStore.AllowSizeGrip = True
    cboStore.AutoDropDown = True
    cboStore.SelForeColor = vbRed
    cboStore.Columns.Add "Stores"
    cboStore.Columns.Add "Seesafe"
    cboStore.Columns(0).Width = 190
    cboStore.Columns(1).Width = 0
    cboStore.BackColorLock = Me.BackColor
    cboStore.EndUpdate
    
End Sub


Private Sub cmdAll_Click()
    If cboStore.Items.ItemCount = 0 Then Exit Sub
    lngTPID = 0
    cboStore.Items.SelectItem(cboStore.Items(0)) = False
End Sub

Private Sub cmdClose_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub


Private Sub Form_Initialize()
Dim ar() As String
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate

End Sub

Private Sub Form_Load()
    Set tlStores = New z_TextList
    tlStores.Load ltStores
    SetupPT
    LoadStores
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Width = 5500
    Height = 4100
    bCancelled = False
End Sub

Private Sub LoadStores()
Dim vntItem As Variant
    On Error GoTo errHandler
Dim i As Integer
Dim ar() As String
    If tlStores.Count = 0 Then Exit Sub
    cboStore.BeginUpdate
    ReDim ar(0 To 1, tlStores.Count)
    cboStore.Items.RemoveAllItems
    i = 1
        For Each vntItem In tlStores
            ar(0, i) = tlStores.Item(i + 1)
            ar(1, i) = tlStores.Key(i + 1)
            i = i + 1
        Next
    cboStore.PutItems ar
    cboStore.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStoresTA.LoadStores"
End Sub

Private Sub cboProductType_SelectionChanged()
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    strPT = cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0)
End Sub
Private Sub cboStore_SelectionChanged()
    lngStoreID = tlStores.Key(cboStore.Items.CellCaption(cboStore.Items.SelectedItem, 0))
    strStoreName = cboStore.Items.CellCaption(cboStore.Items.SelectedItem, 0)
End Sub

Property Get StoreID() As Long
    StoreID = lngStoreID
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
Property Get StoreName() As String
    StoreName = strStoreName
End Property
Property Get PTName() As String
    PTName = strPT
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
