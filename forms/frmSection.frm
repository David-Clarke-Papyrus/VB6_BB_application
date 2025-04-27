VERSION 5.00
Begin VB.Form frmSection 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Allocation product to section"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3735
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1125
   End
   Begin VB.ComboBox cboSection 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1380
      TabIndex        =   2
      Top             =   630
      Width           =   2490
   End
   Begin VB.CommandButton cmdAddSection 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      Height          =   315
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   750
   End
   Begin VB.TextBox txtSection 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1395
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   3870
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   645
      Width           =   1080
   End
End
Attribute VB_Name = "frmSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim oPOL As a_POL
Dim flgLoading As Boolean

Public Sub Component(pPOL As a_POL)
    Set oPOL = pPOL
End Sub
Private Sub cboSection_Click()
    If flgLoading Then Exit Sub
    oPOL.SetSection cboSection
    txtSection = oPOL.Section
End Sub
Private Sub cmdAddSection_Click()
    oPOL.SetSection cboSection
    txtSection = oPOL.Section
End Sub

Private Sub cmdClose_Click()
    If oPOL.Section > "" Then
       Unload Me
    Else
        MsgBox "You must allocate the product to a section.", vbExclamation + vbOKOnly, "Validation"
    End If
End Sub

Private Sub Form_Load()
    flgLoading = True
    txtSection = oPOL.Section
    LoadCombo cboSection, oPC.Configuration.Sections
    flgLoading = False
End Sub

Private Sub txtSection_Validate(Cancel As Boolean)
    oPOL.SetSectionAll txtSection
    txtSection = oPOL.Section
End Sub

