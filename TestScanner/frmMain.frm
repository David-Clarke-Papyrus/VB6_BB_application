VERSION 5.00
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6150
      Top             =   2145
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Serial status"
      Height          =   570
      Left            =   315
      TabIndex        =   7
      Top             =   2535
      Width           =   720
   End
   Begin VB.CommandButton cmdCharCOunt 
      Caption         =   "Chars in Queue"
      Height          =   570
      Left            =   300
      TabIndex        =   6
      Top             =   1875
      Width           =   720
   End
   Begin VB.CommandButton cmdCLose 
      Caption         =   "Close"
      Height          =   570
      Left            =   2790
      TabIndex        =   5
      Top             =   1185
      Width           =   720
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   570
      Left            =   1980
      TabIndex        =   4
      Top             =   1200
      Width           =   720
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   570
      Left            =   1125
      TabIndex        =   3
      Top             =   1200
      Width           =   720
   End
   Begin IOLib.IO IO1 
      Left            =   4125
      Top             =   1380
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin VB.TextBox txtIn 
      Height          =   930
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   225
      Width           =   4185
   End
   Begin VB.TextBox txtOut 
      Height          =   1050
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3225
      Width           =   4155
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Go"
      Height          =   570
      Left            =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdCharCOunt_Click()
    txtOut = IO1.NumCharsInQue
End Sub

Private Sub cmdCLose_Click()
    txtOut = ""
    txtOut = IO1.Close
End Sub

Private Sub cmdList_Click()
    txtOut = ""
    For i = 0 To 10
    txtOut = txtOut & IO1.ListPorts(i, 1) + " "
    Next i
End Sub

Private Sub cmdOpen_Click()
    txtOut = IO1.Open(txtIn & ":", "baud=9600 parity=N data=8 stop=1") 'Open a serial Port.
    IO1.wr
End Sub

Private Sub cmdRead_Click()
    txtOut = IO1.ReadString(10)
End Sub

Private Sub cmdStatus_Click()
' txtOut = (IO1.SerialStatus = SERIAL_CTS_TXHOLD)
 'MsgBox CStr(SERIAL_CTS_TXHOLD)
End Sub

