VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOETAChange 
   Caption         =   "Change ETA date for P.O."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCo 
      Caption         =   "Change"
      Height          =   480
      Left            =   1485
      TabIndex        =   4
      Top             =   1965
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker dtpETA 
      Height          =   390
      Left            =   1365
      TabIndex        =   1
      Top             =   1260
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   688
      _Version        =   393216
      Format          =   62914561
      CurrentDate     =   40697
   End
   Begin VB.TextBox txtPOCode 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   450
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "New ETA date"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   975
      TabIndex        =   3
      Top             =   975
      Width           =   2475
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "P.O.Code"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   930
      TabIndex        =   2
      Top             =   225
      Width           =   2475
   End
End
Attribute VB_Name = "frmPOETAChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCo_Click()
Dim oSQL As New z_SQL

    If MsgBox("You are changing the ETA date for subscription order: " & txtPOCode & " to " & Format(dtpETA, "DD/MM/YYYY"), vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If

    oSQL.RunProc "ChangePOETA", Array(dtpETA, Me.txtPOCode), ""
    MsgBox "ETA changed"
    Unload Me
End Sub

Private Sub Form_Load()
    dtpETA.Value = Date
End Sub
