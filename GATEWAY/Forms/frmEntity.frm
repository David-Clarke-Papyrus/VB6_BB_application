VERSION 5.00
Begin VB.Form frmEntity 
   Caption         =   "Select entity to report"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   3135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2100
      Width           =   915
   End
   Begin VB.ListBox lbEntity 
      Height          =   1425
      Left            =   450
      TabIndex        =   0
      Top             =   345
      Width           =   2115
   End
End
Attribute VB_Name = "frmEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
Dim arOP As New arOP
Dim rs As ADODB.Recordset
Dim oRet As New z_Retrieval

    Set rs = oRet.GetOS(lbEntity)
    arOP.Component rs
    arOP.Show vbModal

End Sub

Private Sub Form_Load()
    LoadListboxSimple lbEntity, oPC.Entities
End Sub
