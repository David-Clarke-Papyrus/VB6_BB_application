VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Settings"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWRD 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show MS-WORD in progress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2055
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2070
      Width           =   1035
   End
   Begin VB.CommandButton cmdFT 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
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
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   750
      Width           =   3705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Catalogue template"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   195
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

