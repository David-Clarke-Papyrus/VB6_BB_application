VERSION 5.00
Begin VB.Form frmShowError 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Error"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   Icon            =   "frmShowError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtError 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   15
      Width           =   12300
   End
End
Attribute VB_Name = "frmShowError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001. All rights reserved.
'version 3.00, 07/15/2001
'Simple window to show error report
'=========================================================================================
Option Explicit

Public Property Let ErrorReport(ByVal AReport As String)
    txtError.text = AReport
    If Not Visible Then ShowSelf
    On Error Resume Next
    Me.SetFocus
End Property

'Try to show non-modal. If there is already modal window in application,
'then it will fail, and we'll show as modal.
Private Sub ShowSelf()
    On Error Resume Next
    Show
    If Err.Number <> 0 Then Show vbModal
End Sub

Private Function ShowNormal()
End Function

Private Sub Form_Resize()
  txtError.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

