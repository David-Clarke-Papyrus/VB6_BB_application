VERSION 5.00
Begin VB.Form frmIDCustomer 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Product requires identification"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCounterfoil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1830
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1950
      Width           =   2700
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2940
      Width           =   1260
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1305
      MaxLength       =   100
      TabIndex        =   0
      Top             =   825
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "and the counterfoil number (if applicable)"
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
      Left            =   315
      TabIndex        =   4
      Top             =   1500
      Width           =   5700
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter customer name for this product . . ."
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
      Left            =   345
      TabIndex        =   3
      Top             =   360
      Width           =   5700
   End
End
Attribute VB_Name = "frmIDCustomer"
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
Public Sub component(pname As String)
    If pname > "" Then
        txtCode = pname
        txtCode.Locked = True
    End If
    
End Sub

Private Sub Form_Load()
    If Len(txtCode) > 0 Then
        txtCounterfoil.SetFocus
    End If
End Sub

Private Sub txtCode_Change()
    strCustomername = txtCode
End Sub
Private Sub txtCounterfoil_Change()
    strCounterfoil = txtCounterfoil
End Sub
Public Property Get CustomerName() As String
    CustomerName = strCustomername
End Property
Public Property Get Counterfoil() As String
    Counterfoil = strCounterfoil
End Property
Private Sub txtCode_GotFocus()
    AutoSelect Controls("txtCode")
End Sub
Private Sub txtCounterfoil_GotFocus()
    AutoSelect Controls("txtCounterfoil")
End Sub

