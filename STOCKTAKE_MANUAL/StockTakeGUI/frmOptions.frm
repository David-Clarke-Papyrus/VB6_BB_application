VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Optiona"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1245
      TabIndex        =   4
      Top             =   2190
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   360
      TabIndex        =   3
      Top             =   2190
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "COM port"
      Height          =   1425
      Left            =   345
      TabIndex        =   0
      Top             =   330
      Width           =   1485
      Begin VB.OptionButton Opt2 
         Caption         =   "Com2"
         Height          =   270
         Left            =   225
         TabIndex        =   2
         Top             =   825
         Width           =   1140
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Com1"
         Height          =   480
         Left            =   225
         TabIndex        =   1
         Top             =   225
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintComPort As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If opt1 = True Then
        mintComPort = 1
        SaveSetting App.Title, "Options", "ComPort", "1"
    Else
        mintComPort = 2
        SaveSetting App.Title, "Options", "ComPort", "2"
    End If
    Me.Hide
End Sub


Private Sub Form_Load()
    mintComPort = GetSetting(App.Title, "Options", "ComPort", 1)
    Select Case mintComPort
    Case 1
        Me.opt1 = True
        Opt2 = False
    Case 2
        opt1 = False
        Opt2 = True
    End Select
End Sub
Public Property Get Comport() As Integer
    Comport = mintComPort
End Property
