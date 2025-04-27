VERSION 5.00
Begin VB.Form frmLabelDesign 
   Caption         =   "Mail label design"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrintWidth 
      Height          =   285
      Left            =   2460
      TabIndex        =   11
      Text            =   "110"
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtTopMargin 
      Height          =   285
      Left            =   2460
      TabIndex        =   9
      Text            =   "110"
      Top             =   1275
      Width           =   585
   End
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   2460
      TabIndex        =   5
      Text            =   "250"
      Top             =   300
      Width           =   585
   End
   Begin VB.TextBox txtRowHeight 
      Height          =   285
      Left            =   2460
      TabIndex        =   4
      Text            =   "2050"
      Top             =   615
      Width           =   570
   End
   Begin VB.TextBox txtColumnSpacing 
      Height          =   285
      Left            =   2460
      TabIndex        =   3
      Text            =   "110"
      Top             =   930
      Width           =   585
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   285
      Left            =   1665
      TabIndex        =   2
      Text            =   "Label type 1"
      Top             =   2385
      Width           =   1425
   End
   Begin VB.CommandButton cmdLoadLabelSettings 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   855
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2370
      Width           =   765
   End
   Begin VB.CommandButton cmdSaveLabelsettings 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2370
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Page width"
      Height          =   285
      Left            =   1140
      TabIndex        =   12
      Top             =   1665
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Top margin"
      Height          =   285
      Left            =   1140
      TabIndex        =   10
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      Height          =   285
      Left            =   1155
      TabIndex        =   8
      Top             =   315
      Width           =   570
   End
   Begin VB.Label Label_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Row height"
      Height          =   285
      Left            =   1155
      TabIndex        =   7
      Top             =   660
      Width           =   960
   End
   Begin VB.Label Label_3 
      BackStyle       =   0  'Transparent
      Caption         =   "Column spacing"
      Height          =   285
      Left            =   1140
      TabIndex        =   6
      Top             =   975
      Width           =   1305
   End
End
Attribute VB_Name = "frmLabelDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mDescription As String
Dim mLeft As Long
Dim mRowHeight As Long
Dim mColumnSpacing As Long
Dim mPrintWidth As Long
Dim mTopMargin As Long


Private Sub cmdLoadLabelSettings_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL
Dim Res As Long
    Res = oSQL.RunGetRecordset("SELECT * FROM tMAILLABEL WHERE ML_DESCRIPTION = '" & CLARG(txtSaveAs) & "'", enText, Array(), "", rs)
    If Not rs.eof Then
        mDescription = FNS(rs.fields(1))
        mLeft = FNN(rs.fields(2))
        mRowHeight = FNN(rs.fields(3))
        mColumnSpacing = FNN(rs.fields(4))
        mTopMargin = FNN(rs.fields(5))
        mPrintWidth = FNN(rs.fields(6))
        txtLeft = CStr(mLeft)
        txtRowHeight = CStr(mRowHeight)
        txtColumnSpacing = CStr(mColumnSpacing)
        txtTopMargin = CStr(mTopMargin)
        txtPrintWidth = CStr(mPrintWidth)
        txtSaveAs = mDescription
    Else
        MsgBox "There is no such mail label"
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.cmdLoadLabelSettings_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSaveLabelsettings_Click()
    On Error GoTo errHandler
Dim oZSQL As New z_SQL
    If IsNumeric(txtLeft) Then mLeft = CLng(txtLeft)
    If IsNumeric(txtRowHeight) Then mRowHeight = CLng(txtRowHeight)
    If IsNumeric(txtColumnSpacing) Then mColumnSpacing = CLng(txtColumnSpacing)
    If IsNumeric(txtTopMargin) Then mTopMargin = CLng(txtTopMargin)
    If IsNumeric(txtPrintWidth) Then mPrintWidth = CLng(txtPrintWidth)
    mDescription = CStr(txtSaveAs)

    oZSQL.SaveMailLabel mDescription, mLeft, mRowHeight, mColumnSpacing, mTopMargin, mPrintWidth
    SaveSetting "CENTRAL", "LABELS", "LABELNAME", mDescription
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.cmdSaveLabelsettings_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    mDescription = GetSetting("CENTRAL", "LABELS", "LABELNAME", "")
    If mDescription > "" Then
        txtSaveAs = mDescription
        cmdLoadLabelSettings_Click
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLeft_Change()
    On Error GoTo errHandler
    If IsNumeric(txtLeft) Then mLeft = CLng(txtLeft)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtLeft_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLeft_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtLeft)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtLeft_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRowHeight_Change()
    On Error GoTo errHandler
    If IsNumeric(txtRowHeight) Then mRowHeight = CLng(txtRowHeight)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtRowHeight_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRowHeight_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtRowHeight)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtRowHeight_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtColumnSpacing_Change()
    On Error GoTo errHandler
    If IsNumeric(txtColumnSpacing) Then mColumnSpacing = CLng(txtColumnSpacing)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtColumnSpacing_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtColumnSpacing_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtColumnSpacing)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtColumnSpacing_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTopMargin_Change()
    On Error GoTo errHandler
    If IsNumeric(txtTopMargin) Then mTopMargin = CLng(txtTopMargin)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtTopMargin_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTopMargin_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtTopMargin)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtTopMargin_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrintWidth_Change()
    On Error GoTo errHandler
    If IsNumeric(txtPrintWidth) Then mPrintWidth = CLng(txtPrintWidth)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtPrintWidth_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrintWidth_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtPrintWidth)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtPrintWidth_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Public Property Get LabelLeft() As Long
    LabelLeft = mLeft
End Property
Public Property Get LabelRowHeight() As Long
    LabelRowHeight = mRowHeight
End Property
Public Property Get LabelColumnSpacing() As Long
    LabelColumnSpacing = mColumnSpacing
End Property
Public Property Get LabelTopMargin() As Long
    LabelTopMargin = mTopMargin
End Property
Public Property Get LabelPrintWidth() As Long
    LabelPrintWidth = mPrintWidth
End Property

Private Sub txtSaveAs_Change()
    On Error GoTo errHandler
    mDescription = txtSaveAs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLabelDesign.txtSaveAs_Change", , EA_NORERAISE
    HandleError
End Sub
