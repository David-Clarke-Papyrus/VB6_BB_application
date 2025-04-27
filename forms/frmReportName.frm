VERSION 5.00
Begin VB.Form frmReportName 
   Caption         =   "Report name"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4260
   LinkTopic       =   "Form2"
   ScaleHeight     =   1470
   ScaleWidth      =   4260
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   780
      Width           =   900
   End
   Begin VB.TextBox txtReportname 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmReportName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
End Sub
Public Property Get Reportname() As String
    Reportname = Me.txtReportname
End Property

Private Sub txtReportname_Change()
    Me.cmdOK.Enabled = (Len(txtReportname) > 0)
End Sub
