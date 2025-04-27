VERSION 5.00
Begin VB.Form frmPOSHELP 
   BackColor       =   &H00DFDED2&
   Caption         =   "POS command line help"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00DFDED2&
      Height          =   4545
      Left            =   90
      TabIndex        =   1
      Top             =   195
      Width           =   6465
      Begin VB.TextBox txtHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDED2&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715248&
         Height          =   3975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmPOSHELP.frx":0000
         Top             =   285
         Width           =   5940
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00CDCFAD&
      Cancel          =   -1  'True
      Caption         =   "&Close  (Esc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4890
      Width           =   1560
   End
End
Attribute VB_Name = "frmPOSHELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim str As String
    str = "..           go back a step" & vbCrLf _
        & ".C           cash payment" & vbCrLf _
        & ".CC          credit card payment" & vbCrLf _
        & ".Q           cheque payment" & vbCrLf _
        & ".V           voucher payment" & vbCrLf _
        & ".Dn          delete line number n" & vbCrLf _
        & ".DPn         delete payment line number n" & vbCrLf _
        & ".CN          credit note tendered" & vbCrLf _
        & ".L           change operator" & vbCrLf _
        & "VR           Void and replace a transaction" & vbCrLf _
        & ".Z           cash up and close application"
        
    txtHelp = str
        
End Sub
