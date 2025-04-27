VERSION 5.00
Begin VB.Form frmMissingTitleAuthor 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Missing title or author"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1005
      Width           =   690
   End
   Begin VB.TextBox txtAuthor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   795
      TabIndex        =   2
      Top             =   615
      Width           =   3675
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   795
      TabIndex        =   0
      Top             =   225
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the missing title and/or author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   315
      Left            =   780
      TabIndex        =   5
      Top             =   1035
      Width           =   2910
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   675
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmMissingTitleAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mTitle As String
Dim mAuthor As String
Dim mPID As String
Dim oSQL As New z_SQL

Public Sub component(PID As String, pTitle As String, pAuthor As String)
    mPID = PID
    mTitle = pTitle
    mAuthor = pAuthor
End Sub

Private Sub cmdOK_Click()
    oSQL.RunSQL "update Tproduct set P_Title = '" & Replace(mTitle, "'", "''") & "',P_MainAuthor = '" & mAuthor & "' where P_ID = '" & mPID & "'"
    Me.Hide
End Sub

Private Sub Form_Load()

    txtTitle = mTitle
    txtTitle.Enabled = (Len(txtTitle) = 0)
    txtAuthor = mAuthor
    txtAuthor.Enabled = (Len(txtAuthor) = 0)
    cmdOK.Enabled = Len(txtAuthor) > 0 And Len(txtTitle) > 0
    
End Sub

Private Sub txtAuthor_Change()
    mAuthor = txtAuthor
    cmdOK.Enabled = Len(txtAuthor) > 0 And Len(txtTitle) > 0
End Sub

Private Sub txtTitle_Change()
    mTitle = txtTitle
    cmdOK.Enabled = Len(txtAuthor) > 0 And Len(txtTitle) > 0
End Sub

'Private Sub txtAuthor_Validate(Cancel As Boolean)
'End Sub

'Private Sub txtTitle_Validate(Cancel As Boolean)
'End Sub

Public Property Get Title() As String
    Title = mTitle
End Property
Public Property Get Author() As String
    Author = mAuthor
End Property
