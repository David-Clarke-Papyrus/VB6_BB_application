VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStartConversation 
      Caption         =   "start conversation"
      Height          =   630
      Left            =   1470
      TabIndex        =   1
      Top             =   780
      Width           =   1740
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send message"
      Height          =   630
      Left            =   1485
      TabIndex        =   0
      Top             =   1725
      Width           =   1740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cs As String
Dim cnn As New ADODB.Connection
Dim Msg As String
Dim comstr As String


Private Sub OpenDB()
    On Error GoTo Err_Handler
    
        cs = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=PBKS-SVR;Initial Catalog=SBC;Password=car"
        cnn.Open cs
    Exit Sub
Err_Handler:
    MsgBox Error
End Sub

Private Sub cmdSend_Click()
    cnn.Execute "EXEC SendBrokerMessage"
End Sub

Private Sub cmdStartConversation_Click()
    cnn.Execute "EXEC StartConversation"
End Sub

Private Sub Form_Load()
    OpenDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cnn.Close
End Sub
