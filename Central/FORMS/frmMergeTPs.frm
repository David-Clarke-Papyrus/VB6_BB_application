VERSION 5.00
Begin VB.Form frmMergeTPs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge trading partners"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeep 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   900
      Width           =   555
   End
   Begin VB.CommandButton cmdLose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   930
      Width           =   555
   End
   Begin VB.CommandButton cmdMerge 
      BackColor       =   &H00C4BCA4&
      Caption         =   "MERGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2175
      Width           =   1695
   End
   Begin VB.TextBox txtLose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Top             =   915
      Width           =   2190
   End
   Begin VB.TextBox txtKeep 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4290
      TabIndex        =   0
      Top             =   900
      Width           =   2190
   End
   Begin VB.Label lblKeep 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4275
      TabIndex        =   5
      Top             =   1350
      Width           =   3150
   End
   Begin VB.Label lblLose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   630
      TabIndex        =   4
      Top             =   1350
      Width           =   3150
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this customer . . .    into . . .        this customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1065
      TabIndex        =   2
      Top             =   585
      Width           =   5745
   End
End
Attribute VB_Name = "frmMergeTPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngLoseTPID As Long
Dim lngKeepTPID As Long

Private Sub cmdFetch_Click()
End Sub

Private Sub cmdKeep_Click()
Dim frmC As frmBrowseCustomers2

        Set frmC = New frmBrowseCustomers2
        frmC.Show vbModal
        txtKeep = frmC.Accnum
        lblKeep.Caption = frmC.CustomerName
        lngKeepTPID = frmC.CustomerID
        Unload frmC
        Set frmC = Nothing

End Sub

Private Sub cmdLose_Click()
Dim frmC As frmBrowseCustomers2

        Set frmC = New frmBrowseCustomers2
        frmC.Show vbModal
        txtLose = frmC.Accnum
        lblLose.Caption = frmC.CustomerName
        lngLoseTPID = frmC.CustomerID
        Unload frmC
        Set frmC = Nothing
    
End Sub

Private Sub cmdMerge_Click()
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
Dim lngResult As Long
    If MsgBox("You want to merge " & lblLose.Caption & " into " & lblKeep.Caption & "?.", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If lngLoseTPID > 0 And lngKeepTPID > 0 Then
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
          
        cmd.CommandText = "MERGETPs"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        ' Get parameter value and append parameter.
        Set prm = cmd.CreateParameter("@pKeep", adInteger, adParamInput)
        cmd.Parameters.Append prm
        prm.Value = lngKeepTPID
        Set prm = cmd.CreateParameter("@pLose", adInteger, adParamInput)
        cmd.Parameters.Append prm
        prm.Value = lngLoseTPID
        Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        cmd.Execute
        lngResult = cmd.Parameters(2)
        If lngResult <> 0 Then
            MsgBox "The Merge operation was unsuccessful."
        Else
            MsgBox "The Merge has completed"
        End If
            
    Else
        MsgBox "One or other of the trading partner codes codes is invalid"
    End If
    Screen.MousePointer = vbDefault

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub





