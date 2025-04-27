VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "reset"
      Height          =   630
      Left            =   5400
      TabIndex        =   4
      Top             =   1335
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open database"
      Height          =   630
      Left            =   10605
      TabIndex        =   3
      Top             =   600
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   390
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   135
      Width           =   11925
   End
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
    Dim strTmp As String
Dim i As Integer
Dim e As Object
        cs = Text1
        On Error Resume Next
        cnn.ConnectionTimeout = 7
        cnn.Open cs
        i = 1
       For Each e In cnn.Errors
        With e
           strTmp = strTmp & vbCrLf & "ADO Error # " & i & ":"
           strTmp = strTmp & vbCrLf & "   ADO Error   # " & .Number
           strTmp = strTmp & vbCrLf & "   Description   " & .Description
           strTmp = strTmp & vbCrLf & "   Source        " & .Source
           i = i + 1
        End With
       Next

        MsgBox "Result = " & IIf(strTmp = "", "OK", strTmp)
        If cnn.State <> 0 Then cnn.Close
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

Private Sub Command1_Click()
    SaveSetting "PBKS", "VPN", "CS", cs
    OpenDB
End Sub

Private Sub Command2_Click()
    cs = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=PBKS-SVR;Initial Catalog=SBC;Password=car;connectiontimeout = 20"
    Text1 = cs
End Sub

Private Sub Form_Load()
    cs = GetSetting("PBKS", "VPN", "CS", "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=PBKS-SVR;Initial Catalog=SBC;Password=car;connectiontimeout=10")

    Text1 = cs
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' cnn.Close
End Sub
