VERSION 5.00
Begin VB.Form frmMergeTPs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge trading partners"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   3540
      Picture         =   "frmMergeTPs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2700
      Width           =   1000
   End
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
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1590
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
      Left            =   2805
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1605
      Width           =   555
   End
   Begin VB.Frame fr 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Trading partner type"
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
      Height          =   870
      Left            =   1695
      TabIndex        =   5
      Top             =   105
      Width           =   4500
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   540
         Left            =   195
         ScaleHeight     =   480
         ScaleWidth      =   4110
         TabIndex        =   9
         Top             =   255
         Width           =   4170
         Begin VB.OptionButton optSupplier 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Supplier"
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
            Height          =   345
            Left            =   165
            TabIndex        =   11
            Top             =   60
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Customer"
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
            Height          =   345
            Left            =   2100
            TabIndex        =   10
            Top             =   60
            Width           =   1290
         End
      End
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
      Left            =   570
      TabIndex        =   0
      Top             =   1590
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
      Left            =   4200
      TabIndex        =   1
      Top             =   1575
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
      Left            =   4200
      TabIndex        =   4
      Top             =   2025
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
      Left            =   555
      TabIndex        =   3
      Top             =   2025
      Width           =   3150
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this trading partner . . .    into . . .        this trading partner"
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
      Left            =   990
      TabIndex        =   2
      Top             =   1260
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
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeTPs.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdKeep_Click()
    On Error GoTo errHandler
Dim frmC As frmBrowseCustomers2
Dim frmS As frmBrowseSUppliers2

    If Me.optSupplier = True Then
        Set frmS = New frmBrowseSUppliers2
        frmS.Show vbModal
        txtKeep = frmS.Accnum
        lblKeep.Caption = frmS.SupplierName
        lngKeepTPID = frmS.SupplierID
        Unload frmS
        Set frmS = Nothing
    Else
        Set frmC = New frmBrowseCustomers2
        frmC.Show vbModal
        txtKeep = frmC.Accnum
        lblKeep.Caption = frmC.CustomerName
        lngKeepTPID = frmC.CustomerID
        Unload frmC
        Set frmC = Nothing
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeTPs.cmdKeep_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLose_Click()
    On Error GoTo errHandler
Dim frmC As frmBrowseCustomers2
Dim frmS As frmBrowseSUppliers2

    If Me.optSupplier = True Then
        Set frmS = New frmBrowseSUppliers2
        frmS.Show vbModal
        txtLose = frmS.Accnum
        lblLose.Caption = frmS.SupplierName
        lngLoseTPID = frmS.SupplierID
        Unload frmS
        Set frmS = Nothing
    Else
        Set frmC = New frmBrowseCustomers2
        frmC.Show vbModal
        txtLose = frmC.Accnum
        lblLose.Caption = frmC.CustomerName
        lngLoseTPID = frmC.CustomerID
        Unload frmC
        Set frmC = Nothing
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeTPs.cmdLose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMerge_Click()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer

Dim lngResult As Long
    If MsgBox("You want to merge " & lblLose.Caption & " into " & lblKeep.Caption & "?.", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    If lngLoseTPID = lngKeepTPID Then
        MsgBox "You cannot select the same " & IIf(Me.optSupplier = True, "supplier", "customer") & " for both sides of the merge operation.", vbOKOnly + vbInformation, "Can't do this"
        Exit Sub
    End If
        
        

    Screen.MousePointer = vbHourglass
    If lngLoseTPID > 0 And lngKeepTPID > 0 Then
        Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
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
        cmd.execute
        lngResult = cmd.Parameters(2)
        If lngResult <> 0 Then
            MsgBox "The Merge operation was unsuccessful.", , "Status"
        Else
            MsgBox "The Merge has completed", , "Status"
        End If
            
    Else
        MsgBox "One or other of the trading partner codes codes is invalid"
    End If
    Set cmd = Nothing
    Set prm = Nothing
    Screen.MousePointer = vbDefault
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeTPs.cmdMerge_Click", , EA_NORERAISE
    HandleError
End Sub





