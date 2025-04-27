VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Begin VB.Form frmMergeSECs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge product types"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMerge 
      BackColor       =   &H00C4BCA4&
      Caption         =   "MERGE"
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
      Left            =   4350
      Picture         =   "frmMergeSECs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton cmdCountTo 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Count"
      Height          =   390
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   510
      Width           =   810
   End
   Begin VB.CommandButton cmdCountFrom 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Count"
      Height          =   390
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   540
      Width           =   810
   End
   Begin VB.CommandButton cmdDeleteUnusedSECs 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete all unused sections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2520
      Width           =   3045
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboSECFrom 
      Height          =   390
      Left            =   255
      OleObjectBlob   =   "frmMergeSECs.frx":038A
      TabIndex        =   0
      Top             =   525
      Width           =   3135
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboSECTo 
      Height          =   390
      Left            =   5505
      OleObjectBlob   =   "frmMergeSECs.frx":1734
      TabIndex        =   1
      Top             =   510
      Width           =   3135
   End
   Begin VB.Label lblPTTo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   750
      Left            =   5505
      TabIndex        =   6
      Top             =   960
      Width           =   3090
   End
   Begin VB.Label lblPTFrom 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   750
      Left            =   285
      TabIndex        =   5
      Top             =   1050
      Width           =   3090
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "this section "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6105
      TabIndex        =   4
      Top             =   165
      Width           =   2505
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "into . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3780
      TabIndex        =   3
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this section . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   495
      TabIndex        =   2
      Top             =   195
      Width           =   3120
   End
End
Attribute VB_Name = "frmMergeSECs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngSECFrom As Long
Dim lngCntSECFrom As Long
Dim lngSECTo As Long
Dim lngCntSECTo As Long
Dim strPTFrom As String
Dim strPTTo As String
Dim tlFrom As z_TextList

Private Sub SetupCBOs()
    On Error GoTo errHandler
    cboSECFrom.BeginUpdate
    cboSECFrom.WidthList = 190
    cboSECFrom.HeightList = 162
    cboSECFrom.AllowSizeGrip = True
    cboSECFrom.AutoDropDown = True
    cboSECFrom.SelForeColor = vbRed
    cboSECFrom.Columns.Add "Section"
    cboSECFrom.Columns.Add "Seesafe"
    cboSECFrom.Columns(0).Width = 190
    cboSECFrom.Columns(1).Width = 0
    cboSECFrom.BackColorLock = Me.BackColor
  '  cboSECFrom.style = DropDownList
    cboSECFrom.EndUpdate
    
    cboSECTo.BeginUpdate
    cboSECTo.WidthList = 190
    cboSECTo.HeightList = 162
    cboSECTo.AllowSizeGrip = True
    cboSECTo.AutoDropDown = True
    cboSECTo.SelForeColor = vbRed
    cboSECTo.Columns.Add "Section"
    cboSECTo.Columns.Add "Seesafe"
    cboSECTo.Columns(0).Width = 190
    cboSECTo.Columns(1).Width = 0
    cboSECTo.BackColorLock = Me.BackColor
   ' cboSECFrom.style = DropDownList
    cboSECTo.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.SetupCBOs"
End Sub




Private Sub cmdDeleteUnusedSECs_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("You want to delete all unused sectios?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.DeleteUnusedSECs
    ReloadCombos
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.cmdDeleteUnusedSECs_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdMerge_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("You want to reallocate all products that are of product type: " & strPTFrom & " to product type: " & strPTTo & "?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.MergeSECs lngSECFrom, lngSECTo
    Screen.MousePointer = vbDefault
    MsgBox "The Merge has completed"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.cmdMerge_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim ar() As String
Dim arFrom() As String

    Set tlFrom = New z_TextList
    tlFrom.Load ltSectionsAll
    tlFrom.CollectionAsArray arFrom
    oPC.Configuration.Sections.CollectionAsArray ar
    SetupCBOs
    cboSECFrom.BeginUpdate
    cboSECFrom.PutItems arFrom
    cboSECFrom.EndUpdate

    cboSECTo.BeginUpdate
    cboSECTo.PutItems ar
    cboSECTo.EndUpdate

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountFrom_Click()
    On Error GoTo errHandler
Dim tmp As Long
Dim oSM As New z_StockManager

    If cboSECFrom.Items.SelectCount = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    tmp = tlFrom.Key(cboSECFrom.Items.CellCaption(cboSECFrom.Items.SelectedItem, 0))
    lngSECFrom = tmp
    lngCntSECFrom = oSM.CountProductsPerSECID(lngSECFrom)
    lblPTFrom = lngCntSECFrom & " products"
    strPTFrom = cboSECFrom.Items.CellCaption(cboSECFrom.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.cmdCountFrom_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboSECTo_SelectionChanged()
    On Error GoTo errHandler
    lblPTTo = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.cboSECTo_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboSECFROM_SelectionChanged()
    On Error GoTo errHandler
    lblPTFrom = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.cboSECFROM_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountTo_Click()
    On Error GoTo errHandler
Dim tmp As Long
Dim oSM As New z_StockManager
    If cboSECFrom.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = oPC.Configuration.Sections.Key(cboSECTo.Items.CellCaption(cboSECTo.Items.SelectedItem, 0))
    lngSECTo = tmp
    lngCntSECTo = oSM.CountProductsPerSECID(lngSECTo)
    lblPTTo = lngCntSECTo & " products"
    strPTTo = cboSECTo.Items.CellCaption(cboSECTo.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.cmdCountTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ReloadCombos()
    On Error GoTo errHandler
Dim ar() As String
Dim arFrom() As String
    
    Set tlFrom = Nothing
    Set tlFrom = New z_TextList
    tlFrom.Load ltSectionsAll
    tlFrom.CollectionAsArray arFrom

    oPC.Configuration.RefreshSections
    oPC.Configuration.Sections.CollectionAsArray ar
    SetupCBOs
    cboSECFrom.Items.RemoveAllItems
    cboSECFrom.BeginUpdate
    cboSECFrom.PutItems arFrom
    cboSECFrom.EndUpdate

    cboSECTo.Items.RemoveAllItems
    cboSECTo.BeginUpdate
    cboSECTo.PutItems ar
    cboSECTo.EndUpdate
    lblPTFrom.Caption = ""
    lblPTTo.Caption = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeSECs.ReloadCombos"
End Sub
