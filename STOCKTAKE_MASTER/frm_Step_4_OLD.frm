VERSION 5.00
Begin VB.Form frm_Step_4 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 4 - Clearing negative on-hand quantities"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E8E8DD&
      Caption         =   "Clear negative on-hand quantities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3765
      Left            =   765
      TabIndex        =   2
      Top             =   420
      Width           =   5325
      Begin VB.TextBox txtNegAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1350
         TabIndex        =   4
         Top             =   2520
         Width           =   2775
      End
      Begin VB.CommandButton cmdClearNegQtys 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Clear negative O.H.quantities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   690
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3060
         Width           =   4065
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "(e.g. 22-08-2010 10:30 PM)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   285
         TabIndex        =   7
         Top             =   2130
         Width           =   4860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Set negative on-hand quantities to zero before calculating adjustments (recommended). "
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
         Height          =   780
         Left            =   300
         TabIndex        =   6
         Top             =   435
         Width           =   4905
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Use this this date for the adjustment transaction                   ( it must be before the cut-off date/time of the stock-take)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   180
         TabIndex        =   5
         Top             =   1200
         Width           =   5010
      End
   End
   Begin VB.CommandButton cmdPrev_to_3 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4695
      Width           =   840
   End
   Begin VB.CommandButton cmdNext_To_5 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4695
      Width           =   840
   End
End
Attribute VB_Name = "frm_Step_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1


Dim strFilename As String


Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
End Sub

Private Sub cmdNext_To_5_Click()
    Set frm5 = New frm_Step_5
    frm5.Component oSA
    frm5.Show
    Unload Me
End Sub

Private Sub cmdPrev_to_3_Click()
    Set frm3 = New frm_Step_3
    frm3.Component oSA
    frm3.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub cmdClearNegQtys_Click()
Dim dteAdj As Date
    Screen.MousePointer = vbHourglass
    dteAdj = CDate(txtNegAdj)
    oSA.ClearNegativeQtys dteAdj
    Screen.MousePointer = vbDefault
    MsgBox "Negative on-hand quantities set to zero", vbInformation, "Status"
End Sub


Private Sub txtNegAdj_Validate(Cancel As Boolean)
    Cancel = Not IsDate(txtNegAdj)
End Sub
