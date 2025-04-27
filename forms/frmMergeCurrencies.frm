VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Begin VB.Form frmMergeCurrs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge two currencies"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
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
      Picture         =   "frmMergeCurrencies.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1410
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
   Begin VB.CommandButton cmdDeleteUnusedCurrencies 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete all unused currencies"
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
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3045
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboCurrFrom 
      Height          =   390
      Left            =   255
      OleObjectBlob   =   "frmMergeCurrencies.frx":038A
      TabIndex        =   0
      Top             =   525
      Width           =   3135
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboCurrTo 
      Height          =   390
      Left            =   5505
      OleObjectBlob   =   "frmMergeCurrencies.frx":1734
      TabIndex        =   1
      Top             =   510
      Width           =   3135
   End
   Begin VB.Label lblCurrTo 
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
   Begin VB.Label lblCurrFrom 
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
      Caption         =   "this currency"
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
      Caption         =   "Merge this currency . . ."
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
      Left            =   480
      TabIndex        =   2
      Top             =   195
      Width           =   3120
   End
End
Attribute VB_Name = "frmMergeCurrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngCurrFrom As Long
Dim lngCntCurrFrom As Long
Dim lngCurrTo As Long
Dim lngCntCurrTo As Long
Dim strCurrFrom As String
Dim strCurrTo As String
Dim tlFrom As z_TextList

Private Sub SetupCBOs()
    On Error GoTo errHandler
    cboCurrFrom.BeginUpdate
    cboCurrFrom.WidthList = 190
    cboCurrFrom.HeightList = 162
    cboCurrFrom.AllowSizeGrip = True
    cboCurrFrom.AutoDropDown = True
    cboCurrFrom.SelForeColor = vbRed
    cboCurrFrom.Columns.Add "Currency"
    cboCurrFrom.Columns.Add "Seesafe"
    cboCurrFrom.Columns(0).Width = 190
    cboCurrFrom.Columns(1).Width = 0
    cboCurrFrom.BackColorLock = Me.BackColor
  '  cboCurrFrom.style = DropDownList
    cboCurrFrom.EndUpdate
    
    cboCurrTo.BeginUpdate
    cboCurrTo.WidthList = 190
    cboCurrTo.HeightList = 162
    cboCurrTo.AllowSizeGrip = True
    cboCurrTo.AutoDropDown = True
    cboCurrTo.SelForeColor = vbRed
    cboCurrTo.Columns.Add "Currency"
    cboCurrTo.Columns.Add "Seesafe"
    cboCurrTo.Columns(0).Width = 190
    cboCurrTo.Columns(1).Width = 0
    cboCurrTo.BackColorLock = Me.BackColor
   ' cboCurrFrom.style = DropDownList
    cboCurrTo.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.SetupCBOs"
End Sub


Private Sub cmdDeleteUnusedCurrs_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("You want to delete all unused currency types?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.DeleteUnusedCurrs
    ReloadCombos
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.cmdDeleteUnusedCurrs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMerge_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If strCurrFrom = strCurrTo Then
        MsgBox "You cannot merge documents of the same currency type.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If MsgBox("You want to reallocate all documents that are of currency type: " & strCurrFrom & " to currency type: " & strCurrTo & "?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.MergeCurrencies lngCurrFrom, lngCurrTo
    Screen.MousePointer = vbDefault
    MsgBox "The Merge has completed"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.cmdMerge_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim ar() As String
Dim arFrom() As String

    Set tlFrom = New z_TextList
    tlFrom.Load ltCurrency
    tlFrom.CollectionAsArray arFrom
 '   oPC.Configuration.ProductTypes.CollectionAsArray ar
    SetupCBOs
    cboCurrFrom.BeginUpdate
    cboCurrFrom.PutItems arFrom
    cboCurrFrom.EndUpdate

    cboCurrTo.BeginUpdate
    cboCurrTo.PutItems arFrom
    cboCurrTo.EndUpdate

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountFrom_Click()
    On Error GoTo errHandler
Dim tmp As Long
Dim oSQL As New z_SQL
    If cboCurrFrom.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = tlFrom.Key(cboCurrFrom.Items.CellCaption(cboCurrFrom.Items.SelectedItem, 0))
    lngCurrFrom = tmp
    lngCntCurrFrom = oSQL.QtyDocumentsUsingCurrency(lngCurrFrom)
    lblCurrFrom = lngCntCurrFrom & " document" & IIf(lngCntCurrFrom = 1, "", "s")
    strCurrFrom = cboCurrFrom.Items.CellCaption(cboCurrFrom.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.cmdCountFrom_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboCurrTo_SelectionChanged()
    On Error GoTo errHandler
    strCurrTo = cboCurrTo.Items.CellCaption(cboCurrTo.Items.SelectedItem, 0)
    lblCurrTo = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.cboCurrTo_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboCurrFROM_SelectionChanged()
    On Error GoTo errHandler
    strCurrFrom = cboCurrFrom.Items.CellCaption(cboCurrFrom.Items.SelectedItem, 0)
    lblCurrFrom = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.cboCurrFROM_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountTo_Click()
    On Error GoTo errHandler
Dim tmp As Long
Dim oSQL As New z_SQL

If cboCurrTo.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = tlFrom.Key(cboCurrTo.Items.CellCaption(cboCurrTo.Items.SelectedItem, 0))
    lngCurrTo = tmp
    lngCntCurrTo = oSQL.QtyDocumentsUsingCurrency(lngCurrTo)
    lblCurrTo = lngCntCurrTo & " document" & IIf(lngCntCurrTo = 1, "", "s")
    strCurrTo = cboCurrTo.Items.CellCaption(cboCurrTo.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.cmdCountTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ReloadCombos()
    On Error GoTo errHandler
Dim ar() As String
Dim arFrom() As String
    
    Set tlFrom = Nothing
    Set tlFrom = New z_TextList
    tlFrom.Load ltCurrency
    tlFrom.CollectionAsArray arFrom

    oPC.Configuration.RefreshProductTypes
   ''' oPC.Configuration.Currencies.CollectionAsArray ar
    SetupCBOs
    cboCurrFrom.Items.RemoveAllItems
    cboCurrFrom.BeginUpdate
    cboCurrFrom.PutItems arFrom
    cboCurrFrom.EndUpdate

    cboCurrTo.Items.RemoveAllItems
    cboCurrTo.BeginUpdate
    cboCurrTo.PutItems ar
    cboCurrTo.EndUpdate
    lblCurrFrom.Caption = ""
    lblCurrTo.Caption = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeCurrs.ReloadCombos"
End Sub
