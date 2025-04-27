VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dayend procedure"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Caption         =   "Dayend running . . . "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   825
      TabIndex        =   0
      Top             =   1155
      Width           =   4830
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oDE As z_Dayend
Attribute oDE.VB_VarHelpID = -1


Public Sub DoWork()
On Error GoTo errHandler
    
    Set oDE = New z_Dayend
    
    Call oDE.RunDayend(0)
    
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DoWork", , EA_NORERAISE
    HandleErrorQuiet True
    Unload Me
End Sub


Private Sub oDE_Status(pMsg As String, pErr As Boolean)
    lblProgress.Caption = pMsg
    If pErr Then
        LogSaveToFile pMsg
    End If
End Sub

Private Sub oDE_COMPLETE()
    Set oDE = Nothing
    Unload Me
End Sub
