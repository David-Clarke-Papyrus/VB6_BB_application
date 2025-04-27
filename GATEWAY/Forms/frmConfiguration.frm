VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmConfiguration 
   Caption         =   "Configuration"
   ClientHeight    =   6705
   ClientLeft      =   -360
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   240
      TabIndex        =   2
      Top             =   210
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   970
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmConfiguration.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkNielsen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkLScheme"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkStockSharing"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAuditing"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Connection and sales query"
      TabPicture(1)   =   "frmConfiguration.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "txtQ"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "FTP setting"
      TabPicture(2)   =   "frmConfiguration.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblHeadings(0)"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "Label14"
      Tab(2).Control(4)=   "Label15"
      Tab(2).Control(5)=   "lstConnections"
      Tab(2).Control(6)=   "chkFTPPassive"
      Tab(2).Control(7)=   "txtFTPDefaultFolder"
      Tab(2).Control(8)=   "txtFTPPassword"
      Tab(2).Control(9)=   "txtFTPUsername"
      Tab(2).Control(10)=   "txtFTPAddress"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Central FTP Details"
      TabPicture(3)   =   "frmConfiguration.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label10"
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(2)=   "Label3"
      Tab(3).Control(3)=   "Label2"
      Tab(3).Control(4)=   "txtFTPAddress_C"
      Tab(3).Control(5)=   "txtFTPUsername_C"
      Tab(3).Control(6)=   "txtFTPPassword_C"
      Tab(3).Control(7)=   "txtFTPDefaultFolder_C"
      Tab(3).Control(8)=   "chkFTPPassive_C"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Loyalty customers query"
      TabPicture(4)   =   "frmConfiguration.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label11"
      Tab(4).Control(1)=   "txtLCQ"
      Tab(4).ControlCount=   2
      Begin VB.CheckBox chkAuditing 
         Caption         =   "Auditing"
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
         Height          =   480
         Left            =   1050
         TabIndex        =   35
         Top             =   4410
         Width           =   4410
      End
      Begin VB.CheckBox chkStockSharing 
         Caption         =   "Stock sharing among stores active"
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
         Height          =   480
         Left            =   1050
         TabIndex        =   34
         Top             =   3570
         Width           =   4410
      End
      Begin VB.CheckBox chkLScheme 
         Caption         =   "Loyalty scheme active"
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
         Height          =   480
         Left            =   1050
         TabIndex        =   33
         Top             =   2730
         Width           =   3255
      End
      Begin VB.CheckBox chkNielsen 
         Caption         =   "Nielsen sales reporting active"
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
         Height          =   480
         Left            =   1035
         TabIndex        =   32
         Top             =   1890
         Width           =   3015
      End
      Begin VB.TextBox txtQ 
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
         ForeColor       =   &H8000000D&
         Height          =   3645
         Left            =   -74745
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1260
         Width           =   10230
      End
      Begin VB.TextBox txtFTPAddress 
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
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -73935
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1365
         Width           =   3030
      End
      Begin VB.TextBox txtFTPUsername 
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
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -73935
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2190
         Width           =   3030
      End
      Begin VB.TextBox txtFTPPassword 
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
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -73935
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3015
         Width           =   3030
      End
      Begin VB.TextBox txtFTPDefaultFolder 
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
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -73935
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   3840
         Width           =   3030
      End
      Begin VB.CheckBox chkFTPPassive 
         Caption         =   "Use FTP passive"
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
         Height          =   480
         Left            =   -73935
         TabIndex        =   20
         Top             =   4500
         Width           =   2010
      End
      Begin VB.ListBox lstConnections 
         Height          =   645
         Left            =   -69810
         TabIndex        =   19
         Top             =   1365
         Width           =   3900
      End
      Begin VB.CheckBox chkFTPPassive_C 
         Caption         =   "Use FTP passive"
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
         Height          =   465
         Left            =   -73950
         TabIndex        =   14
         Top             =   4485
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.TextBox txtFTPDefaultFolder_C 
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
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   -73950
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.TextBox txtFTPPassword_C 
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
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   -73965
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3015
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.TextBox txtFTPUsername_C 
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
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   -73950
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2175
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.TextBox txtFTPAddress_C 
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
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   -73950
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1410
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.TextBox txtLCQ 
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
         ForeColor       =   &H8000000D&
         Height          =   3645
         Left            =   -74775
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1215
         Visible         =   0   'False
         Width           =   10230
      End
      Begin VB.Label Label1 
         Caption         =   "Selection query for Nielsen sales data"
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
         Height          =   240
         Left            =   -74670
         TabIndex        =   31
         Top             =   1005
         Width           =   3675
      End
      Begin VB.Label Label15 
         Caption         =   "FTP default folder"
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
         Height          =   240
         Left            =   -73920
         TabIndex        =   29
         Top             =   3570
         Width           =   2490
      End
      Begin VB.Label Label14 
         Caption         =   "FTP password"
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
         Height          =   240
         Left            =   -73920
         TabIndex        =   28
         Top             =   2760
         Width           =   2490
      End
      Begin VB.Label Label13 
         Caption         =   "FTP username"
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
         Height          =   240
         Left            =   -73920
         TabIndex        =   27
         Top             =   1950
         Width           =   2490
      End
      Begin VB.Label Label12 
         Caption         =   "FTP address"
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
         Height          =   240
         Left            =   -73935
         TabIndex        =   26
         Top             =   1095
         Width           =   2490
      End
      Begin VB.Label lblHeadings 
         Caption         =   "Dial-up network connections:"
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
         Height          =   210
         Index           =   0
         Left            =   -69810
         TabIndex        =   25
         Top             =   1140
         Width           =   2865
      End
      Begin VB.Label Label2 
         Caption         =   "FTP address"
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
         Height          =   225
         Left            =   -73965
         TabIndex        =   18
         Top             =   1095
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label3 
         Caption         =   "FTP username"
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
         Height          =   225
         Left            =   -73935
         TabIndex        =   17
         Top             =   1950
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label6 
         Caption         =   "FTP password"
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
         Height          =   225
         Left            =   -73935
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label10 
         Caption         =   "FTP default folder"
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
         Height          =   225
         Left            =   -73935
         TabIndex        =   15
         Top             =   3570
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label11 
         Caption         =   "Selection query for loyalty customers"
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
         Height          =   240
         Left            =   -74700
         TabIndex        =   9
         Top             =   960
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   915
   End
   Begin VB.CommandButton cmdCancek 
      BackColor       =   &H00CCC8BB&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "FTP address"
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
      Height          =   240
      Left            =   600
      TabIndex        =   7
      Top             =   1125
      Width           =   2490
   End
   Begin VB.Label Label5 
      Caption         =   "FTP username"
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
      Height          =   240
      Left            =   600
      TabIndex        =   6
      Top             =   1905
      Width           =   2490
   End
   Begin VB.Label Label7 
      Caption         =   "FTP password"
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
      Height          =   240
      Left            =   600
      TabIndex        =   5
      Top             =   2700
      Width           =   2490
   End
   Begin VB.Label Label8 
      Caption         =   "FTP default folder"
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
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   2490
   End
   Begin VB.Label Label9 
      Caption         =   "Folder for grower's files"
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
      Height          =   240
      Left            =   900
      TabIndex        =   3
      Top             =   2745
      Width           =   2490
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 3
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags  As Long
   lpfnCallback  As Long
   lParam As Long
   iImage As Long
End Type
Private WithEvents fInet As wininet
Attribute fInet.VB_VarHelpID = -1



Private Sub cmdCancek_Click()
    On Error GoTo errHandler
    oPC.Configuration.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdCancek_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim strError As String
    oPC.Configuration.ApplyEdit strError
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
Dim strDuns() As String
Dim lngIndex  As Long
    txtQ = oPC.Configuration.Q
    txtFTPAddress = oPC.Configuration.FTPAddress
    txtFTPUsername = oPC.Configuration.FTPUsername
    txtFTPPassword = oPC.Configuration.FTPPassword
    txtFTPDefaultFolder = oPC.Configuration.FTPDefaultFolder
    chkFTPPassive = IIf(oPC.Configuration.FTPPassive, 1, 0)
    txtLCQ = oPC.Configuration.LCQ
'    txtFTPAddress_C = oPC.CentralFTPAddress
'    txtFTPUsername_C = oPC.Configuration.CentralFTPUsername
'    txtFTPPassword_C = oPC.Configuration.CentralFTPPassword
'    txtFTPDefaultFolder_C = oPC.Configuration.CentralFTPDefaultFolder
'    chkFTPPassive_C = IIf(oPC.Configuration.CentralFTPPassive, 1, 0)
    Me.chkNielsen = IIf(oPC.Configuration.NielsenActive, 1, 0)
    Me.chkLScheme = IIf(oPC.Configuration.LoyaltySchemeActive, 1, 0)
    Me.chkStockSharing = IIf(oPC.Configuration.StockSharingACtive, 1, 0)
    If oPC.GetProperty("INTERNETDIALUP") = "TRUE" Then
        Set fInet = New wininet
        Call fInet.ListDUNs(strDuns)
        lstConnections.Clear
        For lngIndex = 0 To UBound(strDuns)
            lstConnections.AddItem strDuns(lngIndex)
            If strDuns(lngIndex) = oPC.Configuration.DUN Then
                lstConnections.Selected(lngIndex) = True
            End If
        Next
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.Form_Load", , EA_NORERAISE
    HandleError
End Sub





Private Sub lstConnections_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.DUN = lstConnections
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lstConnections_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtFTPAddress_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.FTPAddress = FNS(txtFTPAddress)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtFTPAddress_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFTPUsername_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.FTPUsername = FNS(txtFTPUsername)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtFTPUsername_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFTPPassword_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.FTPPassword = FNS(txtFTPPassword)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtFTPPassword_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFTPDefaultFolder_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.FTPDefaultFolder = FNS(txtFTPDefaultFolder)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtFTPDefaultFolder_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkFTPPassive_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.FTPPassive = FNB(chkFTPPassive)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkFTPPassive_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkNielsen_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.NielsenActive = FNB(chkNielsen)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkNielsen_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkLScheme_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.LoyaltySchemeActive = FNB(chkLScheme)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkLScheme_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkStockSharing_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.StockSharingACtive = FNB(chkStockSharing)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkStockSharing_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkAuditing_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.AuditingActive = FNB(chkAuditing)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkAuditing_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtQ_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.Q = FNS(txtQ)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtQ_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub












'Private Sub txtFTPAddress_C_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    oPC.Configuration.CentralFTPAddress = FNS(txtFTPAddress_C)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.txtFTPAddress_C_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtFTPUsername_C_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    oPC.Configuration.CentralFTPUsername = FNS(txtFTPUsername_C)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.txtFTPUsername_C_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtFTPPassword_C_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    oPC.Configuration.CentralFTPPassword = FNS(txtFTPPassword_C)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.txtFTPPassword_C_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtFTPDefaultFolder_C_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    oPC.Configuration.CentralFTPDefaultFolder = FNS(txtFTPDefaultFolder_C)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.txtFTPDefaultFolder_C_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub chkFTPPassive_C_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    oPC.Configuration.CentralFTPPassive = FNB(chkFTPPassive_C)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.chkFTPPassive_C_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'

Private Sub txtLCQ_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.LCQ = FNS(txtLCQ)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtLCQ_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub












Public Function GetDatabaseFolder() As String
    On Error GoTo errHandler
'Opens a Treeview control that displays the directories in a computer
Dim lngpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

    szTitle = "Please select the database connection as it has either not been set or has been moved."
    With tBrowseInfo
       .hWndOwner = 0
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lngpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lngpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lngpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        ' comment buy Urs:
        ' The path name will be saved in z_DatabasePersist.dbConnect only if DB is opened
        ' successfuly, else it will be saved as empty string to force the select path box
        ' to be opened again....
'       SaveSetting App.Title, "Settings", "Databasefolder", sBuffer
        
       GetDatabaseFolder = sBuffer
   End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.GetDatabaseFolder"
End Function

