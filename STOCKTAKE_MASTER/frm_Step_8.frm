VERSION 5.00
Begin VB.Form frm_Step_8 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 8 - Finalize"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFinalize 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Finalize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   900
      Width           =   2610
   End
   Begin VB.CommandButton cmdPrev_to_7 
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
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4545
      Width           =   840
   End
   Begin VB.CommandButton cmdNext_To_9 
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
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4515
      Width           =   840
   End
End
Attribute VB_Name = "frm_Step_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1
Dim strSQL As String
Dim strFilename As String
Dim strTitle As String
Dim dteDateTime As Date

Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
End Sub





Private Sub cmdFinalize_Click()
    If MsgBox("Confirm you wish to finalize the stock-take. Changes will not be possible after this.", vbInformation + vbYesNo, "Confirm") = vbYes Then
        Screen.MousePointer = vbHourglass
        oSA.Finalize
'
'        Me.txtTotalItems = oSA.TotalItems
'        Me.txtTotalProducts = oSA.TotalProducts
'        Me.txtValueOfStockCost = oSA.ValueOfStockCostF
'        Me.txtValueOfStockRetail = oSA.ValueOfStockRetailF
'        Me.txtAvgDisc = oSA.AvgDiscountF
        Screen.MousePointer = vbDefault
        MsgBox "Stock take is finalized.", vbInformation, "Status"
    End If
End Sub

Private Sub cmdNext_To_9_Click()
    Set frm9 = New frm_Step_9
    frm9.Component oSA
    frm9.Show
    Unload Me
End Sub

Private Sub cmdPrev_to_7_Click()
    Set frm7 = New frm_Step_7
    frm7.Component oSA
    frm7.Show
    Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

