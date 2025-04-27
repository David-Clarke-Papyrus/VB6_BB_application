VERSION 5.00
Begin VB.Form frmMainPosUtilities 
   BackColor       =   &H00E0E0E0&
   Caption         =   "POS Client Utilities"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCLearData 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Clear local data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2115
   End
End
Attribute VB_Name = "frmMainPosUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCLearData_Click()
Dim frm As New frmUtilities
    frm.Show vbModal
End Sub
