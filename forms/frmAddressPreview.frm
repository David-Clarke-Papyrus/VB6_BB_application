VERSION 5.00
Begin VB.Form frmAddressPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Address view"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   Icon            =   "frmAddressPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFormailing 
      BackColor       =   &H00D3D3CB&
      Caption         =   "For mailing"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   4170
      TabIndex        =   35
      Top             =   2235
      Width           =   1200
   End
   Begin VB.PictureBox frame3 
      BackColor       =   &H00D3D3CB&
      Height          =   765
      Left            =   5880
      ScaleHeight     =   705
      ScaleWidth      =   1755
      TabIndex        =   32
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton optAir 
         BackColor       =   &H00D3D3CB&
         Caption         =   "air"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   105
         TabIndex        =   34
         Top             =   0
         Width           =   825
      End
      Begin VB.OptionButton optSurface 
         BackColor       =   &H00D3D3CB&
         Caption         =   "surface"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   105
         TabIndex        =   33
         Top             =   390
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6780
      Picture         =   "frmAddressPreview.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3255
      Width           =   1125
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   360
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3465
      Width           =   2850
   End
   Begin VB.CheckBox chkGetsCat 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Gets catalogue"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   4170
      TabIndex        =   28
      Top             =   1980
      Width           =   1650
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Height          =   1575
      Left            =   3645
      TabIndex        =   13
      Top             =   0
      Width           =   4275
      Begin VB.TextBox txtBusphone 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   555
         Width           =   2955
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   255
         Width           =   2955
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   855
         Width           =   2955
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1155
         Width           =   2955
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Phone"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   20
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Phone 2"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   90
         TabIndex        =   19
         Top             =   615
         Width           =   600
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "fax"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   225
         TabIndex        =   18
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Email"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   270
         TabIndex        =   17
         Top             =   1200
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   3300
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2820
         Width           =   1350
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2520
         Width           =   2325
      End
      Begin VB.TextBox txtAddressee 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   420
         Width           =   3300
      End
      Begin VB.TextBox txtL6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2220
         Width           =   2625
      End
      Begin VB.TextBox txtL5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1920
         Width           =   2625
      End
      Begin VB.TextBox txtL4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1620
         Width           =   2625
      End
      Begin VB.TextBox txtL3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   2625
      End
      Begin VB.TextBox txtL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1020
         Width           =   2625
      End
      Begin VB.TextBox txtL1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "PCode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   27
         Top             =   2835
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   180
         TabIndex        =   25
         Top             =   2535
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Addressee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Province"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   45
         TabIndex        =   6
         Top             =   2235
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Town"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   5
         Top             =   1935
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   4
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   2
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   1
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Dispatch mode to this address"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   75
      TabIndex        =   30
      Top             =   3510
      Width           =   2250
   End
End
Attribute VB_Name = "frmAddressPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAdd As a_Address

Sub component(pAdd As a_Address)
    On Error GoTo errHandler
    Set oAdd = pAdd
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddressPreview.component(pAdd)", pAdd
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddressPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 1200
        Left = 250
        Height = 4440
        Width = 8150
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddressPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Sub LoadControls()
    On Error GoTo errHandler
    Me.txtDescription = oAdd.Description
    Me.txtAddressee = oAdd.Addressee
    Me.txtL1 = oAdd.Line1
    Me.txtL2 = oAdd.Line2
    Me.txtL3 = oAdd.Line3
    Me.txtL4 = oAdd.Line4
    Me.txtL5 = oAdd.Line5
    Me.txtL6 = oAdd.Line6
    Me.txtPCode = oAdd.pCode
    Me.chkFormailing = IIf(oAdd.ForMailing, 1, 0)
    If oPC.SupportsCatalogue = True Then
        Me.chkGetsCat = IIf(oAdd.GetsCatalogue, 1, 0)
    Else
        Me.chkGetsCat.Visible = False
    End If
    Me.txtPhone = oAdd.Phone
    Me.txtFax = oAdd.Fax
    Me.txtBusphone = oAdd.BusPhone
    Me.txtEmail = oAdd.EMail
    Me.optAir = (oAdd.PostageType = 1)
    Me.optSurface = (oAdd.PostageType = 2)
    Me.txtCountry = oAdd.CountryName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddressPreview.LoadControls"
End Sub


Private Sub Form_DblClick()
    On Error GoTo errHandler

    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText oAdd.AddressMailing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddressPreview.Form_DblClick", , EA_NORERAISE
    HandleError
End Sub

