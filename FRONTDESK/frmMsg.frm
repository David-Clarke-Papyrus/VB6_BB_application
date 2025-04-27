VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1995
   ClientLeft      =   240
   ClientTop       =   240
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   780
      Width           =   3345
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
