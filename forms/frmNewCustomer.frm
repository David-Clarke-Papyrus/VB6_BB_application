VERSION 5.00
Begin VB.Form frmNewCustomer 
   BackColor       =   &H00D3D3CB&
   Caption         =   "New customer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBusiness 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Business customer"
      Height          =   495
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2355
      Width           =   4110
   End
   Begin VB.CommandButton cmdPrivate 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Private customer"
      Height          =   495
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1670
      Width           =   4110
   End
   Begin VB.CommandButton cmdBookclub 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Book club customer"
      Height          =   495
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   985
      Width           =   4110
   End
   Begin VB.CommandButton cmdLoyaltyCustomer 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Loyalty customer"
      Height          =   495
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   300
      Width           =   4110
   End
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBookclub_Click()
    Unload Me
    Forms(0).NewCustomer enBookclub
End Sub

Private Sub cmdBusiness_Click()
    Unload Me
    Forms(0).NewCustomer enBusiness
End Sub

Private Sub cmdLoyaltyCustomer_Click()
    Unload Me
    Forms(0).NewLoyaltyCustomer
End Sub

Private Sub cmdPrivate_Click()
    Unload Me
    Forms(0).NewCustomer enPrivate

End Sub

Private Sub Form_Load()
    cmdBookclub.Visible = oPC.SupportsBookClubsTF
    cmdLoyaltyCustomer.Visible = oPC.SupportsLoyaltyCustomersTF

End Sub
