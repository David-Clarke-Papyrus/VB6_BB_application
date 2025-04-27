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
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3060
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2145
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Height          =   555
      Left            =   3060
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open database connection"
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   450
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMainConnectionString As String
Dim oCnn As New ADODB.Connection
Private Sub Command1_Click()
        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & txtServer & ";Initial Catalog=PBKS;User Id=sa;Password=car;Connect Timeout=45"
        oCnn.Open strMainConnectionString

End Sub

Private Sub Command2_Click()
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strXML As String
Dim strm


    Set cmd = New ADODB.Command
    cmd.CommandText = "GetDeliveryForExportXML"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DOCCODE", adVarChar, adParamInput, 25, Text1)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@XML", adVarChar, adParamOutput, 4000)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oCnn

    Set strm = CreateObject("ADODB.Stream")
    strm.Open
    cmd.Properties("Output Stream") = strm
    cmd.Execute , , adCmdStoredProc Or adExecuteStream
    strm.Position = 0
    strXML = strm.ReadText
    strm.Close
    Set cmd = Nothing
    Set strm = Nothing


    oCnn.Execute "SENDBROKERMESSAGE '" & strXML & "'"
End Sub
