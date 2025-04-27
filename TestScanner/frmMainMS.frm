VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMainMS 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBreakOff 
      Caption         =   "Break ff"
      Height          =   570
      Left            =   2010
      TabIndex        =   9
      Top             =   1875
      Width           =   720
   End
   Begin VB.CommandButton cmdBreak 
      Caption         =   "Break on"
      Height          =   570
      Left            =   1140
      TabIndex        =   8
      Top             =   1875
      Width           =   720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6150
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      SThreshold      =   2
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
      Left            =   5415
      TabIndex        =   5
      Top             =   240
      Width           =   720
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   570
      Left            =   1125
      TabIndex        =   4
      Top             =   2520
      Width           =   720
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   570
      Left            =   4575
      TabIndex        =   3
      Top             =   240
      Width           =   720
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
      Caption         =   "Send"
      Height          =   570
      Left            =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   720
   End
End
Attribute VB_Name = "frmMainMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdBreak_Click()
    MSComm1.Break = True
End Sub

Private Sub cmdBreakOff_Click()
    MSComm1.Break = False
End Sub

Private Sub cmdCLose_Click()
    MSComm1.PortOpen = False
End Sub

Private Sub cmdList_Click()
    MSComm1.Output = CStr(Trim(Me.txtIn))
End Sub

Private Sub cmdOpen_Click()
    MSComm1.PortOpen = True
    MSComm1.DTREnable = True
    txtOut = "CTSHolding = " & MSComm1.CTSHolding & vbCrLf & txtOut
    txtOut = "DSRHolding = " & MSComm1.DSRHolding & vbCrLf & txtOut
End Sub

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
    
        Case comEvRing
            txtOut = "ComEVRing  " & vbCrLf & txtOut
        Case comEvDSR
            txtOut = "comEvDSR  " & vbCrLf & txtOut
        Case comEvCTS
            txtOut = "comEvCTS  " & vbCrLf & txtOut
        Case comEvCD
            txtOut = "comEvCD  " & vbCrLf & txtOut
        Case comEvReceive
            txtOut = "comEvReceive:" & MSComm1.Input & vbCrLf & txtOut
            
        Case comEvSend
            txtOut = "comEvSend  " & vbCrLf & txtOut
        Case Else
                txtOut = MSComm1.CommEvent & vbCrLf & txtOut

    End Select
End Sub
