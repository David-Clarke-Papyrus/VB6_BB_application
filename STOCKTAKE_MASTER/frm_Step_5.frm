VERSION 5.00
Begin VB.Form frm_Step_4 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 4 - Consolidation"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConsolidate 
      BackColor       =   &H00D8D9C4&
      Caption         =   "Consolidate all captured data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2130
      Width           =   3345
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
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4500
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
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
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
Private oBatch As z_SQL

Dim strFilename As String


Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
End Sub



Private Sub cmdConsolidate_Click()
    Screen.MousePointer = vbHourglass
    Set oBatch = New z_SQL
    oBatch.RunProc "sp_STOCKTAKE_CONSOLIDATE", Array(), ""
    Set oBatch = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Consolidation complete", vbInformation, "Status"
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


