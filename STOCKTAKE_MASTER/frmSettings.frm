VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWRD 
      Caption         =   "Show MS-WORD in progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   180
      TabIndex        =   5
      Top             =   1410
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3510
      TabIndex        =   4
      Top             =   2430
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2460
      TabIndex        =   3
      Top             =   2430
      Width           =   1035
   End
   Begin VB.CommandButton cmdFT 
      Caption         =   ">>>"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   780
      Width           =   585
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   180
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTemplate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   750
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "Catalogue template"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   2085
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTemplate As String
Dim flgShowWORD As Boolean

Friend Sub Component(pstrTemplate As String, pShowWord As Boolean)
    strTemplate = pstrTemplate
    Me.txtTemplate = strTemplate
    Me.chkWRD.Value = IIf(pShowWord, 1, 0)
    Me.Refresh
End Sub

Private Sub chkWRD_Click()
    flgShowWORD = chkWRD.Value
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFT_Click()
    CD1.ShowOpen
    strTemplate = CD1.FileName
    Me.txtTemplate = strTemplate
End Sub

Private Sub cmdOK_Click()
    SaveSetting App.Title, "Settings", "Template", strTemplate
    SaveSetting App.Title, "Settings", "ShowWORD", flgShowWORD
    Me.Hide
End Sub
Friend Function Template() As String
    Template = strTemplate
End Function
Friend Function ShowWORD() As Boolean
    ShowWORD = flgShowWORD
End Function

