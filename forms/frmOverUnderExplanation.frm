VERSION 5.00
Begin VB.Form frmOverUnderExplanation 
   Caption         =   "Explanation"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4680
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
      Height          =   585
      Left            =   1695
      Picture         =   "frmOverUnderExplanation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1635
      Width           =   1110
   End
   Begin VB.TextBox txtExplanation 
      ForeColor       =   &H8000000D&
      Height          =   1365
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   4245
   End
End
Attribute VB_Name = "frmOverUnderExplanation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub component(txt As String)
    On Error GoTo errHandler
    Me.txtExplanation = txt
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOverUnderExplanation.component(txt)", txt
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOverUnderExplanation.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get Explanation() As String
    On Error GoTo errHandler
    Explanation = Trim(txtExplanation)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOverUnderExplanation.Explanation"
End Property
