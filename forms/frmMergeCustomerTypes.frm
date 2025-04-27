VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Begin VB.Form frmMergeCustomerTypes 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge customer types"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMerge 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Merge"
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
      Left            =   4215
      Picture         =   "frmMergeCustomerTypes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1575
      Width           =   1000
   End
   Begin VB.CommandButton cmdDeleteUnusedPTs 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete all unused customer types"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2355
      Width           =   3195
   End
   Begin VB.CommandButton cmdCountFrom 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Count"
      Height          =   390
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   765
      Width           =   810
   End
   Begin VB.CommandButton cmdCountTo 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Count"
      Height          =   390
      Left            =   8565
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   735
      Width           =   810
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboCTFrom 
      Height          =   390
      Left            =   135
      OleObjectBlob   =   "frmMergeCustomerTypes.frx":038A
      TabIndex        =   0
      Top             =   750
      Width           =   3135
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboCTTo 
      Height          =   390
      Left            =   5385
      OleObjectBlob   =   "frmMergeCustomerTypes.frx":1734
      TabIndex        =   1
      Top             =   735
      Width           =   3135
   End
   Begin VB.Label lblCTFrom 
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
      Left            =   165
      TabIndex        =   10
      Top             =   1290
      Width           =   3090
   End
   Begin VB.Label lblCTTo 
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
      Left            =   5385
      TabIndex        =   9
      Top             =   1200
      Width           =   3090
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
      Left            =   3990
      TabIndex        =   8
      Top             =   390
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this customer type . . ."
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
      Left            =   375
      TabIndex        =   6
      Top             =   420
      Width           =   3120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "this customer type "
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
      Left            =   5985
      TabIndex        =   5
      Top             =   390
      Width           =   2505
   End
End
Attribute VB_Name = "frmMergeCustomerTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCTFrom As String
Dim strCTTO As String
Dim lngCTFrom As Long
Dim lngCTTo As Long
Dim tlFrom As z_TextList
Dim lngCntCTFrom As Long
Dim lngCntCTTo As Long

Private Sub SetupCBOs()
    On Error GoTo errHandler
    cboCTFrom.BeginUpdate
    cboCTFrom.WidthList = 190
    cboCTFrom.HeightList = 162
    cboCTFrom.AllowSizeGrip = True
    cboCTFrom.AutoDropDown = True
    cboCTFrom.SelForeColor = vbRed
    cboCTFrom.Columns.Add "Customer type"
    cboCTFrom.Columns.Add "Seesafe"
    cboCTFrom.Columns(0).Width = 190
    cboCTFrom.Columns(1).Width = 0
    cboCTFrom.BackColorLock = Me.BackColor
    cboCTFrom.EndUpdate
    
    cboCTTo.BeginUpdate
    cboCTTo.WidthList = 190
    cboCTTo.HeightList = 162
    cboCTTo.AllowSizeGrip = True
    cboCTTo.AutoDropDown = True
    cboCTTo.SelForeColor = vbRed
    cboCTTo.Columns.Add "Customer type"
    cboCTTo.Columns.Add "Seesafe"
    cboCTTo.Columns(0).Width = 190
    cboCTTo.Columns(1).Width = 0
    cboCTTo.BackColorLock = Me.BackColor
    cboCTTo.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.SetupCBOs"
End Sub

Private Sub cmdDeleteUnusedPTs_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("You want to delete all unused customer types from the dictionary?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.DeleteUnusedCTs
    ReloadCombos
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.cmdDeleteUnusedPTs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMerge_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("You want to reallocate all customers that are of customer type: " & strCTFrom & " to customer type: " & strCTTO & "?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lngCTFrom = tlFrom.Key(cboCTFrom.Items.CellCaption(cboCTFrom.Items.SelectedItem, 0))
    lngCTTo = tlFrom.Key(cboCTTo.Items.CellCaption(cboCTTo.Items.SelectedItem, 0))
    oSM.MergeCTs lngCTFrom, lngCTTo
    Screen.MousePointer = vbDefault
    MsgBox "The merge has completed"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.cmdMerge_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim ar() As String
Dim arFrom() As String

    Set tlFrom = New z_TextList
    tlFrom.Load ltCustomerTypeAllExceptLoyalty
    tlFrom.CollectionAsArray arFrom
    SetupCBOs
    cboCTFrom.BeginUpdate
    cboCTFrom.PutItems arFrom
    cboCTFrom.EndUpdate

    cboCTTo.BeginUpdate
    cboCTTo.PutItems arFrom
    cboCTTo.EndUpdate

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountFrom_Click()
    On Error GoTo errHandler
Dim tmp As Long
Dim oSM As New z_StockManager
    If cboCTFrom.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = tlFrom.Key(cboCTFrom.Items.CellCaption(cboCTFrom.Items.SelectedItem, 0))
   ' If tmp = lngPTFrom Then Exit Sub
    lngCTFrom = tmp
    lngCntCTFrom = oSM.CountCustomersPerCTID(lngCTFrom)
    lblCTFrom = lngCntCTFrom & " customers"
    strCTFrom = cboCTFrom.Items.CellCaption(cboCTFrom.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.cmdCountFrom_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboCTTo_SelectionChanged()
    On Error GoTo errHandler
    lblCTTo = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.cboCTTo_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboCTFrom_SelectionChanged()
    On Error GoTo errHandler
    lblCTFrom = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.cboCTFrom_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountTo_Click()
    On Error GoTo errHandler
Dim tmp As Long
Dim oSM As New z_StockManager
    If cboCTTo.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = tlFrom.Key(cboCTTo.Items.CellCaption(cboCTTo.Items.SelectedItem, 0))
    lngCTTo = tmp
    lngCntCTTo = oSM.CountCustomersPerCTID(lngCTTo)
    lblCTTo = lngCntCTTo & " customers"
    strCTTO = cboCTTo.Items.CellCaption(cboCTTo.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.cmdCountTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ReloadCombos()
    On Error GoTo errHandler
Dim ar() As String
Dim arFrom() As String
    
    Set tlFrom = Nothing
    Set tlFrom = New z_TextList
    tlFrom.Load ltCustomerTypeAllExceptLoyalty
    tlFrom.CollectionAsArray arFrom

    SetupCBOs
    cboCTFrom.Items.RemoveAllItems
    cboCTFrom.BeginUpdate
    cboCTFrom.PutItems arFrom
    cboCTFrom.EndUpdate

    cboCTTo.Items.RemoveAllItems
    cboCTTo.BeginUpdate
    cboCTTo.PutItems arFrom
    cboCTTo.EndUpdate
    lblCTFrom.Caption = ""
    lblCTTo.Caption = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.ReloadCombos"
End Sub

Private Sub lblPTFrom_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCustomerTypes.lblPTFrom_Click", , EA_NORERAISE
    HandleError
End Sub
