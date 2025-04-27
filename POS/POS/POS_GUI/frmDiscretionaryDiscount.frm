VERSION 5.00
Begin VB.Form frmDiscretionaryDiscount 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Discretionary discount"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1410
      Width           =   1260
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1050
      MaxLength       =   10
      TabIndex        =   0
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Discretionary discount"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   255
      Width           =   3765
   End
End
Attribute VB_Name = "frmDiscretionaryDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strCustomername As String
Dim strCounterfoil As String

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Public Property Get DiscountRate() As Double
    DiscountRate = CDbl(Trim(txtDiscount))
End Property

Private Sub txtDiscount_Validate(Cancel As Boolean)
Dim bOK As Boolean
Dim dblDisc As Double

    bOK = IsNumeric(txtDiscount)
    If bOK Then
        dblDisc = CDbl(txtDiscount)
        If dblDisc > 100 Or dblDisc < 0 Then
            bOK = False
        End If
    End If
    Cancel = Not bOK
            
End Sub
