VERSION 5.00
Begin VB.Form frmPeriodSelection 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Select period to post to"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3240
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmPeriodSelection.frx":0000
      Left            =   1050
      List            =   "frmPeriodSelection.frx":002B
      TabIndex        =   1
      Text            =   "1"
      Top             =   270
      Width           =   765
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Continue"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   900
      Picture         =   "frmPeriodSelection.frx":005A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   900
      Width           =   1000
   End
End
Attribute VB_Name = "frmPeriodSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngPeriod As Long

Private Sub OKButton_Click()
    lngPeriod = Me.Combo1
    Me.Hide
End Sub
Public Property Get Period()
    Period = lngPeriod
End Property
