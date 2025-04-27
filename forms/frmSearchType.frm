VERSION 5.00
Begin VB.Form frmSearchType 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Search style"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAntiquarian 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Antiquarian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   630
      Width           =   1515
   End
   Begin VB.CommandButton cmdNormal 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Normal"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   1515
   End
End
Attribute VB_Name = "frmSearchType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String


Private Sub cmdNormal_Click()
    On Error GoTo errHandler
    strType = "N"
    SaveSetting "PBKS", "Startup", "SearchStyle", "N"
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSearchType.cmdNormal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAntiquarian_Click()
    On Error GoTo errHandler
    strType = "A"
    SaveSetting "PBKS", "Startup", "SearchStyle", "A"
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSearchType.cmdAntiquarian_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get SearchType() As String
    SearchType = strType
End Property

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strSetting As String
    strSetting = GetSetting("PBKS", "Startup", "SearchStyle", "N")
    If strSetting = "N" Then
        Me.cmdNormal.Default = True
    Else
        Me.cmdAntiquarian.Default = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSearchType.Form_Load", , EA_NORERAISE
    HandleError
End Sub
