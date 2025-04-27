VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmScannerSettings 
   Caption         =   "Scanner settings"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   Icon            =   "frmScannerSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
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
      Left            =   375
      TabIndex        =   9
      Top             =   2175
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3825
      TabIndex        =   0
      Top             =   2175
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   345
      TabIndex        =   6
      Top             =   285
      Width           =   4700
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   2861
         TabIndex        =   5
         Top             =   400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtComPort"
         BuddyDispid     =   196614
         OrigLeft        =   2860
         OrigTop         =   400
         OrigRight       =   3100
         OrigBottom      =   760
         Max             =   16
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboBaudRate 
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
         ItemData        =   "frmScannerSettings.frx":038A
         Left            =   1995
         List            =   "frmScannerSettings.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1000
         Width           =   2300
      End
      Begin VB.TextBox txtComPort 
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
         Left            =   1980
         TabIndex        =   2
         Text            =   "1"
         Top             =   400
         Width           =   1100
      End
      Begin VB.Label lblBaudRate 
         AutoSize        =   -1  'True
         Caption         =   "&Baud rate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   3
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label lblComPort 
         AutoSize        =   -1  'True
         Caption         =   "&COM port :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   1
         Top             =   465
         Width           =   930
      End
   End
   Begin VB.Label lblTimeout 
      Height          =   225
      Left            =   3060
      TabIndex        =   8
      Top             =   2205
      Width           =   1965
   End
   Begin VB.Label lblRecord 
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1965
   End
End
Attribute VB_Name = "frmScannerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "PBKS", "ManualCount", "ScannerComPort", txtComPort
    SaveSetting "PBKS", "ManualCount", "BaudRate", cboBaudRate
    Unload Me
End Sub

Private Sub Form_Load()
    cboBaudRate.AddItem "115200"
    cboBaudRate.AddItem "38400"
    cboBaudRate.ListIndex = 0
    
    txtComPort = GetSetting("PBKS", "ManualCount", "ScannerComPort", "1")
    cboBaudRate.Text = GetSetting("PBKS", "ManualCount", "BaudRate", "115200")

End Sub


