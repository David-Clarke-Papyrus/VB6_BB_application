VERSION 5.00
Begin VB.Form frmEditCatHead 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Catalogue heading"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
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
      Height          =   615
      Left            =   7455
      Picture         =   "frmEditCatHead.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2745
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   8400
      Picture         =   "frmEditCatHead.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2745
      Width           =   1000
   End
   Begin VB.ComboBox cboBelongsTo 
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
      Height          =   360
      ItemData        =   "frmEditCatHead.frx":0714
      Left            =   45
      List            =   "frmEditCatHead.frx":0716
      TabIndex        =   4
      Top             =   1935
      Width           =   9345
   End
   Begin VB.TextBox txtSortTag 
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
      Height          =   405
      Left            =   825
      TabIndex        =   2
      Top             =   1005
      Width           =   1245
   End
   Begin VB.TextBox txtDescription 
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
      Height          =   405
      Left            =   15
      TabIndex        =   0
      Top             =   420
      Width           =   9330
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Belongs to"
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
      Height          =   300
      Left            =   75
      TabIndex        =   5
      Top             =   1620
      Width           =   1620
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort tag"
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
      Height          =   300
      Left            =   45
      TabIndex        =   3
      Top             =   1050
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmEditCatHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCH As a_cathead
Dim flgLoading As Boolean
Dim tlHeadings As z_TextList
Dim bCancel As Boolean

Public Sub component(pCATHEAD As a_cathead)
    On Error GoTo errHandler
    Set oCH = pCATHEAD
    oCH.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.component(pCATHEAD)", pCATHEAD
End Sub

Private Sub cboBelongsTo_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If cboBelongsTo <> "<None>" Then
        oCH.Parent = tlHeadings.key(Me.cboBelongsTo)
        oCH.parentHeading = cboBelongsTo
    Else
        oCH.Parent = 0
        oCH.parentHeading = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.cboBelongsTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oCH.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    oCH.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Set tlHeadings = New z_TextList
    tlHeadings.Load ltCatalogueHeadings
    Me.txtDescription = oCH.Description
    Me.txtSortTag = oCH.SortTag
    LoadCombo Me.cboBelongsTo, tlHeadings
    cboBelongsTo.AddItem "<None>", 0
    If oCH.parentHeading = "" Then
        cboBelongsTo = "<None>"
    Else
        Me.cboBelongsTo = oCH.parentHeading
    End If
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDescription_Change()
            On Error Resume Next
Dim intPos As Integer
    On Error Resume Next
    bCancel = Not oCH.SetDescription(txtDescription)
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oCH.Description
      txtDescription.SelStart = intPos - 1
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.txtDescription_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDescription = oCH.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.txtDescription_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDescription_Validate(Cancel As Boolean)
            On Error Resume Next
    Cancel = bCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSortTag_Change()
            On Error Resume Next
Dim intPos As Integer
    On Error Resume Next
   bCancel = Not oCH.SetSortTag(txtSortTag)
    If Err Then
      Beep
      intPos = txtSortTag.SelStart
      txtSortTag = oCH.SortTag
      txtSortTag.SelStart = intPos - 1
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.txtSortTag_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSortTag_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtSortTag = oCH.SortTag
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.txtSortTag_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSortTag_Validate(Cancel As Boolean)
            On Error Resume Next
    Cancel = bCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEditCatHead.txtSortTag_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
