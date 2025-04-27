VERSION 5.00
Begin VB.Form frmTiming 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Timing"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   2760
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmTiming.frx":0000
      Top             =   285
      Width           =   4380
   End
End
Attribute VB_Name = "frmTiming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub component(val As String)
    Me.Text1 = val
End Sub
