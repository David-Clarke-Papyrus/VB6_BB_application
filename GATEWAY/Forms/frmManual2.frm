VERSION 5.00
Begin VB.Form frmManual2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Export-Import"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Start POS Server polling"
      Height          =   375
      Left            =   6375
      TabIndex        =   16
      Top             =   6615
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop POS Server polling"
      Height          =   375
      Left            =   6375
      TabIndex        =   15
      Top             =   6195
      Width           =   2745
   End
   Begin VB.CommandButton cmdUpdateStockSharing 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Update stock"
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
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5355
      Width           =   3675
   End
   Begin VB.CommandButton cmdSSGo 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1110
      Width           =   660
   End
   Begin VB.TextBox txtLastSSExport 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6360
      TabIndex        =   11
      Top             =   1080
      Width           =   2205
   End
   Begin VB.CommandButton cmdLCGo 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1185
      Width           =   660
   End
   Begin VB.TextBox txtLastLoyaltySent 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1125
      TabIndex        =   8
      Top             =   1155
      Width           =   2205
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Delete Confirmed Stock Exports"
      Enabled         =   0   'False
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
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4470
      Width           =   3675
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Import Stock Confirmations"
      Enabled         =   0   'False
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
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3675
      Width           =   3675
   End
   Begin VB.CommandButton cmdExportStock 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Export Stock file"
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
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   3675
   End
   Begin VB.CommandButton cmdCreateStockFile 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Create Stock file"
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
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2085
      Width           =   3675
   End
   Begin VB.CommandButton cmdDeleteConfirmed 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Delete Confirmed Exports"
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
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4485
      Width           =   3675
   End
   Begin VB.CommandButton cmdImportLoyaltyConfirmations 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Import Loyalty Confirmations"
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
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3690
      Width           =   3675
   End
   Begin VB.CommandButton cmdExportLoyalty 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Export Loyalty files"
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
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2895
      Width           =   3675
   End
   Begin VB.CommandButton cmdCreateLoyalty 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Create Loyalty files"
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
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last stock export"
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
      Height          =   390
      Left            =   6345
      TabIndex        =   12
      Top             =   705
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last loyalty export"
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
      Height          =   390
      Left            =   1110
      TabIndex        =   9
      Top             =   780
      Width           =   2220
   End
End
Attribute VB_Name = "frmManual2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private Sub cmdCreateLoyalty_Click()
'    On Error GoTo errHandler
'Dim oLC As New z_Loyalty
'    oLC.Component oPC
'    oLC.CreateLoyaltyExtractionFile
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmManual2.cmdCreateLoyalty_Click", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub cmdCreateStockFile_Click()
'    On Error GoTo errHandler
'Dim oLC As New z_Loyalty
'    oLC.Component oPC
'    oLC.CreateStockSharingExtractionFile
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmManual2.cmdCreateStockFile_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cmdDeleteConfirmed_Click()
    On Error GoTo errHandler
Dim oEx As New z_Export
    oEx.DeleteReceipted
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.cmdDeleteConfirmed_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExportLoyalty_Click()
    On Error GoTo errHandler
Dim oEx As New z_Export
    oEx.SendLoyalty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.cmdExportLoyalty_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExportStock_Click()
    On Error GoTo errHandler
Dim oEx As New z_Export
    oEx.SendStockSharing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.cmdExportStock_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdImportLoyaltyConfirmations_Click()
'Dim oEX As New z_Export
'    oEX.FetchLCConfirmations
'End Sub

Private Sub cmdLCGo_Click()
    On Error GoTo errHandler
oPC.COShort.Execute "Update tNielsen Set N_LastDateLCSent = '" & ReverseDate(Me.txtLastLoyaltySent) & "'"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.cmdLCGo_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSalesExport_Click()
    On Error GoTo errHandler
    Dim oEx As New z_Export
  '  oEx.SendNielsenFiles
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.cmdSalesExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSSGo_Click()
    On Error GoTo errHandler
oPC.COShort.Execute "Update tNielsen Set N_LastLoyaltyExportDate = '" & ReverseDate(Me.txtLastLoyaltySent) & "'"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.cmdSSGo_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdUpdateStockSharing_Click()
'    On Error GoTo errHandler
'Dim oLC As New z_Loyalty
'    oLC.Component oPC
'    oLC.UpdateStock
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmManual2.cmdUpdateStockSharing_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler
'ControlPOSServerPolling False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.Command1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Command2_Click()
    On Error GoTo errHandler
'ControlPOSServerPolling True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.Command2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLastLoyaltySent_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not (IsDate(txtLastLoyaltySent))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.txtLastLoyaltySent_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtLastSSExport_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not (IsDate(txtLastSSExport))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual2.txtLastSSExport_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
