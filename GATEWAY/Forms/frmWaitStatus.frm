VERSION 5.00
Begin VB.Form frmWaitStatus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl1 
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
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   255
      Width           =   2550
   End
End
Attribute VB_Name = "frmWaitStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Component(pMsg As String)
    Me.lbl1.Caption = pMsg
    If Len(lbl1.Caption) > 15 Then
        Me.Width = 6500
        lbl1.Width = 5800
    End If
    
End Sub
