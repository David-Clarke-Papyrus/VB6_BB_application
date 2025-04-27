VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDiagnostics 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Diagnostics"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7541
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmDiagnostics.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDB"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Missing exchange numbers"
      TabPicture(1)   =   "frmDiagnostics.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtStationName"
      Tab(1).Control(1)=   "cmdCheck"
      Tab(1).Control(2)=   "txtMissing"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtStationName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   -74520
         TabIndex        =   5
         Top             =   720
         Width           =   1965
      End
      Begin VB.CommandButton cmdCheck 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Run check"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -72420
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txtMissing 
         Height          =   2745
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1290
         Width           =   1905
      End
      Begin VB.Label Label2 
         Caption         =   "Station name (till point)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   -74550
         TabIndex        =   6
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label lblDB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   3150
         Left            =   180
         TabIndex        =   2
         Top             =   810
         Width           =   6720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Database and connections"
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
         Height          =   285
         Left            =   195
         TabIndex        =   1
         Top             =   510
         Width           =   3210
      End
   End
End
Attribute VB_Name = "frmDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCheck_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim strResult As String

    oSM.GetMissingExchangeNumbers Trim(txtStationName), strResult
    txtMissing = strResult
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDiagnostics.cmdCheck_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Me.lblDB.Caption = "Connection string: " & vbCrLf & oPC.ConnectionString & vbCrLf & _
                        "Server name: " & vbCrLf & oPC.servername & vbCrLf & _
                        "Shared server folder: " & vbCrLf & oPC.SharedFolderRoot
                        
                        

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDiagnostics.Form_Load", , EA_NORERAISE
    HandleError
End Sub

