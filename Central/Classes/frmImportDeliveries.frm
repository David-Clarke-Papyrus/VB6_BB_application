VERSION 5.00
Begin VB.Form frmImportDeliveries 
   Caption         =   "Bulk delivery import from FTP"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdHB 
      Caption         =   "Import from Jonathan Ball"
      Height          =   495
      Left            =   855
      TabIndex        =   2
      Top             =   1785
      Width           =   2685
   End
   Begin VB.CommandButton cmdOnTHeDot 
      Caption         =   "Import from On The Dot"
      Height          =   495
      Left            =   870
      TabIndex        =   1
      Top             =   1125
      Width           =   2685
   End
   Begin VB.CommandButton cmdBooksite 
      Caption         =   "Import from Booksite"
      Height          =   495
      Left            =   870
      TabIndex        =   0
      Top             =   495
      Width           =   2685
   End
End
Attribute VB_Name = "frmImportDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOnTHeDot_Click()
Dim f As New frmDeliveryImport

    f.Component "OTD"
    f.Show
    
End Sub
