VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWarnings 
   Caption         =   "Scanner settings"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   Icon            =   "frmWarnings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D5D5C1&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Canc&el"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   315
      TabIndex        =   4
      Top             =   195
      Width           =   4700
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   3496
         TabIndex        =   3
         Top             =   510
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtQty"
         BuddyDispid     =   196616
         OrigLeft        =   3795
         OrigTop         =   510
         OrigRight       =   4035
         OrigBottom      =   870
         Max             =   16
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   3090
         TabIndex        =   2
         Text            =   "1"
         Top             =   510
         Width           =   405
      End
      Begin VB.Label lblComPort 
         AutoSize        =   -1  'True
         Caption         =   "Warn for quantites in excess of:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   555
         Width           =   2745
      End
   End
   Begin VB.Label lblTimeout 
      Height          =   225
      Left            =   3060
      TabIndex        =   6
      Top             =   2205
      Width           =   1965
   End
   Begin VB.Label lblRecord 
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1965
   End
End
Attribute VB_Name = "frmWarnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "PBKS", "ManualCount", "QtyWarningLimit", txtQty
    Unload Me
End Sub

Private Sub Form_Load()
    
    txtQty = GetSetting("PBKS", "ManualCount", "QtyWarningLimit", "10")

End Sub
