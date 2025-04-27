VERSION 5.00
Begin VB.Form frmPrintRemindersheet 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Print reminder sheet"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1215
      Picture         =   "frmPrintRemindersheet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1635
      Width           =   1000
   End
   Begin VB.CheckBox chkPagePerSupplier 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Separate page per supplier"
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
      Height          =   240
      Left            =   510
      TabIndex        =   0
      Top             =   675
      Value           =   1  'Checked
      Width           =   2910
   End
End
Attribute VB_Name = "frmPrintRemindersheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOrderedOnly As Boolean
Dim strSequence As String


Public Property Get Sequence() As String
    On Error GoTo errHandler
    Sequence = strSequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintRemindersheet.Sequence"
End Property

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintRemindersheet.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
