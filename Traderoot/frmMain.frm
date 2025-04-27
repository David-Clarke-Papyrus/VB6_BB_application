VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Test bed for Trade route merchant"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSignOff 
      Caption         =   "Sign off"
      Height          =   510
      Left            =   4530
      TabIndex        =   3
      Top             =   3135
      Width           =   2325
   End
   Begin VB.CommandButton cmdSignOn 
      Caption         =   "Sign on"
      Height          =   510
      Left            =   4530
      TabIndex        =   1
      Top             =   930
      Width           =   2325
   End
   Begin VB.TextBox txtIn 
      Height          =   420
      Left            =   555
      TabIndex        =   0
      Text            =   "txtIn"
      Top             =   285
      Width           =   6240
   End
   Begin VB.Label lblResults 
      BorderStyle     =   1  'Fixed Single
      Height          =   2670
      Left            =   450
      TabIndex        =   2
      Top             =   4560
      Width           =   6825
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

