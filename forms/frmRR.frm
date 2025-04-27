VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRR 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Rounding rules"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   6600
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
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
      Left            =   5205
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5085
      Width           =   930
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
      Height          =   570
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5085
      Width           =   930
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Enabled         =   0   'False
      Height          =   1410
      Left            =   195
      TabIndex        =   7
      Top             =   3510
      Width           =   6090
      Begin VB.TextBox txtLB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   525
         Width           =   1695
      End
      Begin VB.TextBox txtUB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1950
         TabIndex        =   2
         Top             =   525
         Width           =   1695
      End
      Begin VB.TextBox txtRT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4815
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   405
         Width           =   930
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Figures represent the smallest denomination of currency (e.g. cents)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   165
         TabIndex        =   11
         Top             =   1020
         Width           =   5880
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lower bound"
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
         Height          =   330
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Upper bound"
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
         Height          =   330
         Left            =   1935
         TabIndex        =   9
         Top             =   255
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Round to"
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
         Height          =   330
         Left            =   3720
         TabIndex        =   8
         Top             =   255
         Width           =   1665
      End
   End
   Begin MSComctlLib.ListView lvwRR 
      Height          =   3150
      Left            =   195
      TabIndex        =   0
      Top             =   345
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5556
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lower bound"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Upper bound"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Round to"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEditm 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit selected"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete selected"
      End
   End
End
Attribute VB_Name = "frmRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim chRR As ch_RoundingRule
Dim oRR As a_RoundingRule
Dim flgLoading As Boolean

Private Sub cmdAdd_Click()
    On Error GoTo errHandler
    oRR.ApplyEdit
    oPC.Configuration.RoundingRules.ApplyEdit
    LoadList
    oPC.Configuration.RoundingRules.BeginEdit
    Set oRR = oPC.Configuration.RoundingRules.Add
 '   oRR.BeginEdit
    txtLB = ""
    txtUB = ""
    txtRT = ""
    Frame1.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oRR.CancelEdit
    oPC.Configuration.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNew_Click()
    On Error GoTo errHandler
    If Not oRR Is Nothing Then
        If oRR.IsEditing Then
            oRR.CancelEdit
        End If
    End If
    Set oRR = oPC.Configuration.RoundingRules.Add
    oRR.BeginEdit
    Frame1.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.cmdNew_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim strStatus As String

    oRR.CancelEdit
    oPC.Configuration.ApplyEdit strStatus
    Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
       Left = 50
        TOP = 50
        Height = 6500
        Width = 6830
    End If
    flgLoading = True
    oPC.Configuration.BeginEdit
    Set oRR = oPC.Configuration.RoundingRules.Add
  '  oRR.BeginEdit
    LoadList
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadList()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwRR.ListItems.Clear
    For i = 1 To oPC.Configuration.RoundingRules.Count
        Set objItm = Me.lvwRR.ListItems.Add
        With objItm
            .Key = oPC.Configuration.RoundingRules(i).Key
            .text = oPC.Configuration.RoundingRules(i).LowerBound
            .SubItems(1) = oPC.Configuration.RoundingRules(i).UpperBound
            .SubItems(2) = oPC.Configuration.RoundingRules(i).RoundTo
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.LoadList"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oRR.IsEditing Then oRR.CancelEdit
    If oPC.Configuration.IsEditing Then oPC.Configuration.CancelEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwRR_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.lvwRR_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwRR_DblClick()
    On Error GoTo errHandler
mnuEdit_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.lvwRR_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDelete_Click()
    On Error GoTo errHandler
    oPC.Configuration.RoundingRules.Remove (lvwRR.SelectedItem.Key)
    LoadList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.mnuDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuEdit_Click()
    On Error GoTo errHandler
    If Not oRR Is Nothing Then
        If oRR.IsEditing Then
            oRR.CancelEdit
        End If
    End If

    Set oRR = oPC.Configuration.RoundingRules(lvwRR.SelectedItem.Key)
    txtLB = oRR.LowerBound
    txtUB = oRR.UpperBound
    txtRT = oRR.RoundTo
    oRR.BeginEdit
    Frame1.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.mnuEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLB_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtLB
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtLB_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLB_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtLB = oRR.LowerBound
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtLB_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLB_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oRR.SetLowerBound txtLB
    If Err Then
      Beep
      intPos = txtLB.SelStart
      txtLB = oRR.LowerBoundF
      txtLB.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtLB_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLB_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oRR.SetLowerBound(txtLB)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtLB_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtUB_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtUB
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtUB_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtUB_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtUB = oRR.UpperBound
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtUB_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtUB_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oRR.SetUpperBound txtUB
    If Err Then
      Beep
      intPos = txtUB.SelStart
      txtUB = oRR.UpperBoundF
      txtUB.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtUB_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtUB_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oRR.SetUpperBound(txtUB)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtUB_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRT_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtRT
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtRT_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRT_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtRT = oRR.RoundTo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtRT_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRT_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oRR.SetRoundTo txtRT
    If Err Then
      Beep
      intPos = txtRT.SelStart
      txtRT = oRR.RoundToF
      txtRT.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtRT_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRT_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oRR.SetRoundTo(txtRT)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRR.txtRT_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

