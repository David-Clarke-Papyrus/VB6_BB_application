VERSION 5.00
Begin VB.Form frmHelpGen 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Help"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00CEC7AE&
      Caption         =   "&Close"
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
      Left            =   3570
      Picture         =   "frmHelpGen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1000
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
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
      Height          =   2850
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHelpGen.frx":038A
      Top             =   90
      Width           =   4500
   End
End
Attribute VB_Name = "frmHelpGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub component(pMsg As String, pTitle As String, pWidth As Long, pHeight As Long)
    On Error GoTo errHandler
    txtMsg = pMsg
    Caption = pTitle
    Width = pWidth
    Height = pHeight
    txtMsg.Width = Width - 200
    txtMsg.Height = Height - 600
    txtMsg.Left = (Width - txtMsg.Width) / 2
    cmdclose.TOP = Height - 1100
    cmdclose.Left = Width - 1500
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHelpGen.component(pMsg,pTitle,pWidth,pHeight)", Array(pMsg, pTitle, pWidth, pHeight)
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHelpGen.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub CLoseForm()
    Unload Me
End Sub
