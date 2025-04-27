VERSION 5.00
Begin VB.Form frmLoyalty 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Loyalty customer"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10095
   Begin VB.CheckBox chkNotifySales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Notify of book sales"
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
      Height          =   450
      Left            =   1110
      TabIndex        =   17
      Top             =   6285
      Width           =   2580
   End
   Begin VB.CheckBox chkNotifyLaunches 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Notify of book launches"
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
      Height          =   450
      Left            =   1110
      TabIndex        =   16
      Top             =   6630
      Width           =   2580
   End
   Begin VB.CheckBox chkNotifyPromotions 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Notify of book promotions"
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
      Height          =   450
      Left            =   1110
      TabIndex        =   15
      Top             =   5940
      Width           =   2580
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   105
      TabIndex        =   32
      Top             =   1080
      Width           =   4785
      Begin VB.TextBox txtAddressee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   4
         Top             =   255
         Width           =   3090
      End
      Begin VB.CommandButton cmdDuplicates 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Check for duplicates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2955
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2610
         Width           =   1800
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         TabIndex        =   14
         Top             =   4170
         Width           =   3120
      End
      Begin VB.TextBox txtCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         TabIndex        =   13
         Top             =   3810
         Width           =   1920
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         TabIndex        =   11
         Top             =   3090
         Width           =   1920
      End
      Begin VB.TextBox txtBusphone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         TabIndex        =   12
         Top             =   3450
         Width           =   1920
      End
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   10
         Top             =   2625
         Width           =   1590
      End
      Begin VB.ComboBox cboCNTRY 
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
         ItemData        =   "frmLoyalty.frx":0000
         Left            =   1260
         List            =   "frmLoyalty.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2190
         Width           =   2370
      End
      Begin VB.TextBox txtLine4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   8
         Top             =   1665
         Width           =   3090
      End
      Begin VB.TextBox txtLine3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   7
         Top             =   1320
         Width           =   3090
      End
      Begin VB.TextBox txtLine2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   6
         Top             =   960
         Width           =   3090
      End
      Begin VB.TextBox txtLine1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   5
         Top             =   615
         Width           =   3090
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Addressee"
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
         Height          =   255
         Left            =   135
         TabIndex        =   44
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Email"
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
         Height          =   270
         Left            =   60
         TabIndex        =   42
         Top             =   4230
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Cell"
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
         Height          =   270
         Left            =   60
         TabIndex        =   41
         Top             =   3855
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Work"
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
         Height          =   270
         Left            =   60
         TabIndex        =   40
         Top             =   3495
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Phone"
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
         Height          =   270
         Left            =   60
         TabIndex        =   39
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label Label15 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Postal code"
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
         Height          =   270
         Left            =   90
         TabIndex        =   38
         Top             =   2655
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Country"
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
         Height          =   270
         Left            =   435
         TabIndex        =   37
         Top             =   2220
         Width           =   690
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
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
         Height          =   255
         Left            =   180
         TabIndex        =   36
         Top             =   1755
         Width           =   870
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
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
         Height          =   255
         Left            =   465
         TabIndex        =   35
         Top             =   1395
         Width           =   600
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 2"
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
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 1"
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
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   630
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Interest group"
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
      Height          =   2025
      Left            =   5430
      TabIndex        =   27
      Top             =   1875
      Width           =   4380
      Begin VB.CommandButton cmdRemoveIG 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1380
         Width           =   1050
      End
      Begin VB.CommandButton cmdAddIG 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Add &group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   330
         Width           =   1305
      End
      Begin VB.ComboBox cboIG 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   375
         Width           =   2745
      End
      Begin VB.ListBox lbIG 
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
         Height          =   990
         Left            =   135
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   795
         Width           =   2700
      End
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5415
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   4215
      Width           =   4350
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5580
      TabIndex        =   2
      Top             =   495
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Picture         =   "frmLoyalty.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   8745
      Picture         =   "frmLoyalty.frx":038E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   1000
   End
   Begin VB.TextBox txtAcno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   5430
      TabIndex        =   3
      Top             =   1350
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   165
      TabIndex        =   0
      Top             =   495
      Width           =   3000
   End
   Begin VB.TextBox txtFN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3405
      TabIndex        =   1
      Top             =   495
      Width           =   1965
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
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
      Height          =   255
      Left            =   5355
      TabIndex        =   31
      Top             =   3990
      Width           =   465
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2460
      X2              =   300
      Y1              =   315
      Y2              =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Left            =   5625
      TabIndex        =   26
      Top             =   225
      Width           =   570
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
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
      Height          =   255
      Left            =   3450
      TabIndex        =   25
      Top             =   225
      Width           =   1815
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   870
      Left            =   5055
      TabIndex        =   24
      Top             =   5715
      Width           =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   255
      Left            =   180
      TabIndex        =   23
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Num."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5775
      TabIndex        =   22
      Top             =   1095
      Width           =   1065
   End
End
Attribute VB_Name = "frmLoyalty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCust As a_Customer
Attribute oCust.VB_VarHelpID = -1
Dim flgLoading As Boolean
Private colClassErrors As Collection
Dim XA As New XArrayDB
Dim strEMail As String
Dim oAdd As a_Address
Dim bAlternativeCustomerSelected As Boolean

Public Property Get EMail() As String
    EMail = strEMail
End Property


Private Sub chkNotifyLaunches_Click()
    On Error GoTo errHandler
    oCust.CustNotifyBookLaunch = (chkNotifyLaunches = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.chkNotifyLaunches_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkNotifySales_Click()
    On Error GoTo errHandler
    oCust.CustNotifyBookSale = (chkNotifySales = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.chkNotifySales_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cboCT_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    oCust.CustomerTypeID = oCust.CustomerTypesActive_tl.Key(cboCT)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cboCT_Click", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub chkTemp_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    oCust.CanBeDeleted = (Me.chkTemp = 1)
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.chkTemp_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cboCNTRY_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not cboCNTRY.ListIndex = -1 Then
        oCust.BillTOAddress.CountryID = oCust.BillTOAddress.Countries.Key(cboCNTRY)
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAddress.cboCNTRY_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cboCNTRY_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddIG_Click()
    On Error GoTo errHandler
Dim oIG As New a_IG
    If flgLoading Then Exit Sub
    If cboIG = "" Then Exit Sub
    Set oIG = oCust.InterestGroups.Add
    oIG.BeginEdit
    oIG.TPID = oCust.ID
    oIG.IGID = oCust.InterestGroupsLoyalty_tl.Key(cboIG)
    oIG.Description = cboIG
    oIG.ApplyEdit
    cboIG.RemoveItem cboIG.ListIndex
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
    LoadTPIGs
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cmdAddIG_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdAddIG_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdDuplicates_Click()
    On Error GoTo errHandler
    oCust.LookforDuplicates
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cmdDuplicates_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdDuplicates_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub chkNotifyPromotions_Click()
    On Error GoTo errHandler
    oCust.CustNotifyBookPromotion = (chkNotifyPromotions = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.chkNotifyPromotions_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveIG_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbIG = "" Then Exit Sub
    oCust.InterestGroups.Remove oCust.InterestGroups.Key(Me.lbIG)
    cboIG.AddItem Me.lbIG
    cboIG.ListIndex = 0
    LoadTPIGs
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cmdRemoveIG_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdRemoveIG_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(pCust As a_Customer)
    On Error GoTo errHandler
Dim oAdd As a_Address
Dim oCC As a_IG
    Set oCust = pCust
    oCust.CustomerTypeID = oPC.Configuration.LoyaltyClubTypeID
    
    Me.Caption = "Loyalty customer (edit): " & oCust.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.Component(pCust)", pCust
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.component(pCust)", pCust
End Sub
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = pOK
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.EnableOK(pOK)", pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.EnableOK(pOK)", pOK
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("Confirm you want to cancel entry of loyalty customer", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
    End If
    oCust.CancelEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim oCC As a_IG
Dim frmP As frmCustomerPreview
    If flgLoading Then Exit Sub
    If oCust.IsNew Then
        bAlternativeCustomerSelected = False
        oCust.LookforDuplicates
        If bAlternativeCustomerSelected Then
            oCust.CancelEdit
            Unload Me
            Exit Sub
        End If
    End If
    If oCust.CustomerIndexClashes = True Then
        MsgBox "This account number has already been used for another customer. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If

    
    If oCust.CustNotifyBookLaunch Then
        Set oCC = Nothing
        Set oCC = oCust.InterestGroups.Add
        oCC.BeginEdit
        oCC.Description = oCust.InterestGroupsAll_tl(oPC.Configuration.IGLaunchID)
        oCC.IGID = oPC.Configuration.IGLaunchID
        oCC.ApplyEdit
    End If
    If oCust.CustNotifyBookPromotion Then
        Set oCC = Nothing
        Set oCC = oCust.InterestGroups.Add
        oCC.BeginEdit
        oCC.IGID = oPC.Configuration.IGPromotionID
        oCC.Description = oCust.InterestGroupsAll_tl(oPC.Configuration.IGPromotionID)
        oCC.ApplyEdit
    End If
    If oCust.CustNotifyBookSale Then
        Set oCC = Nothing
        Set oCC = oCust.InterestGroups.Add
        oCC.BeginEdit
        oCC.IGID = oPC.Configuration.IGSaleID
        oCC.Description = oCust.InterestGroupsAll_tl(oPC.Configuration.IGSaleID)
        oCC.ApplyEdit
    End If
    
    Set oCC = Nothing
    Set oCC = oCust.CustomerTypes.Add
    oCC.BeginEdit
    oCC.IGID = oPC.Configuration.LoyaltyClubTypeID
    oCC.Description = oCust.CustomerTypesALL_tl(oPC.Configuration.LoyaltyClubTypeID)
    oCC.ApplyEdit

    oCust.ApplyEdit lngResult
    If lngResult = 0 Then
        Set frmP = New frmCustomerPreview
        frmP.component oCust
        frmP.Show

        Unload Me
    ElseIf lngResult = 22 Then
        MsgBox "You are trying to save a customer with duplicate values." & vbCrLf & "These are likely to be in the Acc No. field or in the address description fields.", , "Can't save"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cmdOK_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    If oCust.IsNew Then
        Set oAdd = oCust.Addresses.Add
        oAdd.BeginEdit
    End If
    If Me.WindowState <> 2 Then
        Me.TOP = 0
        Me.Left = 50
        Me.Height = 8200
        Me.Width = 11000
    End If
    txtName = oCust.Name
    txtFN = oCust.Initials
    txtAcno = oCust.AcNo
    txtTitle = oCust.Title
    txtNote = oCust.Note
    txtAddressee = oCust.BillTOAddress.Addressee
    txtLine1 = oCust.BillTOAddress.Line1
    txtLine2 = oCust.BillTOAddress.Line2
    txtLine3 = oCust.BillTOAddress.Line5
    txtLine4 = oCust.BillTOAddress.Line6
    
    txtPCode = oCust.BillTOAddress.pCode
    txtPhone = oCust.Phone
    txtBusphone = oCust.BillTOAddress.BusPhone
    txtEmail = oCust.BillTOAddress.EMail
    txtCell = oCust.MOBILE
    chkNotifyLaunches = IIf(oCust.CustNotifyBookLaunch, 1, 0)
    chkNotifyPromotions = IIf(oCust.CustNotifyBookPromotion, 1, 0)
    chkNotifySales = IIf(oCust.CustNotifyBookSale, 1, 0)
    LoadCombo cboCNTRY, oCust.BillTOAddress.Countries
    If oCust.BillTOAddress.CountryID > 0 Then
        cboCNTRY.text = oCust.BillTOAddress.Countries(CStr(oCust.BillTOAddress.CountryID))
    ElseIf oPC.Configuration.LocalCountryID Then
        cboCNTRY.text = oCust.BillTOAddress.Countries(CStr(oPC.Configuration.LocalCountryID))
        oCust.BillTOAddress.CountryID = oPC.Configuration.LocalCountryID
    End If
    LoadIGs
    LoadTPIGs
    RestrictInterestGroups
    oCust.GetStatus
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadIGs()
    On Error GoTo errHandler
    LoadCombo Me.cboIG, oCust.InterestGroupsLoyalty_tl
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.LoadIGs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.LoadIGs"
End Sub
Private Sub RestrictInterestGroups()
    On Error GoTo errHandler
Dim oTPIG As a_IG
Dim i As Integer

    For Each oTPIG In oCust.InterestGroups
        For i = cboIG.ListCount To 1 Step -1
            cboIG.ListIndex = i - 1
            If oTPIG.Description = cboIG Then
                cboIG.RemoveItem cboIG.ListIndex
            End If
        Next
    Next
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.RestrictInterestGroups"
End Sub

Private Sub LoadTPIGs()
    On Error GoTo errHandler
Dim oTPIG As a_IG
    With Me.lbIG
        .Clear
        For Each oTPIG In oCust.InterestGroups
            .AddItem oTPIG.Description   ', oTPIG.Key
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.LoadTPIGs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.LoadTPIGs"
End Sub
'Private Sub lvwAddresses_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub



Private Sub oCust_Valid(strMsg As String)
    On Error GoTo errHandler
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.oCust_Valid(strMsg)", strMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.oCust_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
    HandleError
End Sub


Private Sub ShowDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
Dim frm As frmDuplicateCustomers
Dim tmpCust As a_Customer
    
    Set frm = New frmDuplicateCustomers
    frm.component Me.txtName, pDuplicates
    frm.Show vbModal
    If frm.SelectedCustomer > 0 Then
        Set Forms(0).frmMainCustomerPreview = Nothing
        Set Forms(0).frmMainCustomerPreview = New frmCustomerPreview
        Set tmpCust = New a_Customer
        tmpCust.Load frm.SelectedCustomer
        Forms(0).frmMainCustomerPreview.component tmpCust
        Unload frm
        bAlternativeCustomerSelected = True
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.ShowDuplicates(pDuplicates)", pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.ShowDuplicates(pDuplicates)", pDuplicates
End Sub



Private Sub txtAddressee_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetAddressee txtAddressee
    If Err Then
      Beep
      intPos = txtLine1.SelStart
      txtAddressee = oCust.BillTOAddress.Addressee
      txtAddressee.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtAddressee_Validate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAddressee_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub txtAddressee_GotFocus()
    On Error GoTo errHandler
    txtAddressee = oCust.Title & " " & oCust.Initials & " " & oCust.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAddressee_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLine1_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetLine1 txtLine1
    If Err Then
      Beep
      intPos = txtLine1.SelStart
      txtLine1 = oCust.BillTOAddress.Line1
      txtLine1.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtLine1_Validate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine1_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtLine2_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetLine2 txtLine2
    If Err Then
      Beep
      intPos = txtLine2.SelStart
      txtLine2 = oCust.BillTOAddress.Line2
      txtLine2.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtLine2_Validate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine2_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtLine3_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetLine5 txtLine3
    If Err Then
      Beep
      intPos = txtLine3.SelStart
      txtLine3 = oCust.BillTOAddress.Line5
      txtLine3.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtLine3_Validate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine3_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtPCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetPCode txtPCode
    If Err Then
      Beep
      intPos = txtPCode.SelStart
      txtPCode = oCust.BillTOAddress.pCode
      txtPCode.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtPCode_Validate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtLine4_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetLine6 txtLine4
    If Err Then
      Beep
      intPos = txtLine4.SelStart
      txtLine4 = oCust.BillTOAddress.Line6
      txtLine4.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtLine4_Validate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine4_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtDefaultDiscount_LostFocus()
'    On Error GoTo errHandler
'    txtDefaultDiscount = oCust.DefaultDiscountF
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtDefaultDiscount_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtDefaultDiscount_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oCust.SetDefaultDiscount(txtDefaultDiscount)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtDefaultDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub

'Private Sub txtCell_LostFocus()
'    On Error GoTo errHandler
'    txtCell = oCust.Addresses(1).setCell
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtCell_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtCell_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oCust.setCell(txtCell)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtCell_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtCell_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    oCust.setCell (txtCell)
'    If Err Then
'      Beep
'      intPos = txtCell.SelStart
'      txtCell = oCust.Cell
'      txtCell.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtCell_Change", , EA_NORERAISE
'    HandleError
'End Sub
Private Sub txtEMail_LostFocus()
    On Error GoTo errHandler
    txtEmail = oCust.BillTOAddress.EMail
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtEmail_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtEMail_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEMail_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    Cancel = Not oCust.BillTOAddress.SetEmail(txtEmail)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtEmail_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtEMail_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmail_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetEmail txtEmail
  '  oCust.SetPhone txtEmail    'SetPhone (txtEmail)
    If Err Then
      Beep
      intPos = txtEmail.SelStart
      txtEmail = oCust.BillTOAddress.EMail
      txtEmail.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtEmail_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtEmail_Change", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtBusPhone_LostFocus()
    On Error GoTo errHandler
    On Error Resume Next
    txtBusphone = oCust.BillTOAddress.BusPhone
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtBusPhone_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtBusPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBusPhone_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    txtBusphone = PhoneFormat(txtBusphone, oPC.DefaultAreaCode)
    Cancel = Not oCust.BillTOAddress.SetBusPhone(txtBusphone)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtBusPhone_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtBusPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBusPhone_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetBusPhone txtBusphone
  '  oCust.SetPhone txtBusPhone    'SetPhone (txtBusPhone)
    If Err Then
      Beep
      intPos = txtBusphone.SelStart
      txtBusphone = oCust.BillTOAddress.BusPhone
      txtBusphone.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtBusPhone_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtBusPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCell_LostFocus()
    On Error GoTo errHandler
    txtCell = oCust.MOBILE
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtCell_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtCell_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCell_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtCell = PhoneFormat(txtCell, oPC.DefaultAreaCode)
    Cancel = Not oCust.SetCell(txtCell)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtCell_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtCell_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCell_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.SetCell txtCell
    If Err Then
      Beep
      intPos = txtCell.SelStart
      txtCell = oCust.Phone
      txtCell.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtCell_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtCell_Change", , EA_NORERAISE
    HandleError
End Sub





Private Sub txtPhone_LostFocus()
    On Error GoTo errHandler
    txtPhone = oCust.Phone
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtPhone_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    txtPhone = PhoneFormat(txtPhone, oPC.DefaultAreaCode)
    Cancel = Not oCust.SetPhone(txtPhone)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.BillTOAddress.SetPhone txtPhone
  '  oCust.SetPhone txtPhone    'SetPhone (txtPhone)
    If Err Then
      Beep
      intPos = txtPhone.SelStart
      txtPhone = oCust.Phone
      txtPhone.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtPhone_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
    On Error Resume Next
    txtName = oCust.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtName_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.SetName (txtName)
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oCust.Name
      txtName.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtName_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtName_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    oCust.SetName (txtName)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtName_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtName_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtAcno_LostFocus()
    On Error GoTo errHandler
    txtAcno = oCust.AcNo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtAcno_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAcno_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.SetAcNO (txtAcno)
    If Err Then
      Beep
      intPos = txtAcno.SelStart
      txtAcno = oCust.AcNo
      txtAcno.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtAcno_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAcno_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    'Cancel = Not oCust.SetAcNO(txtAcno)
    On Error Resume Next
    oCust.SetAcNO txtAcno
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtAcno_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAcno_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_LostFocus()
    On Error GoTo errHandler
    txtFN = oCust.Initials
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtFN_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtFN_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.SetInitials (txtFN)
    If Err Then
      Beep
      intPos = txtFN.SelStart
      txtFN = oCust.Initials
      txtFN.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtFN_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtFN_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    Cancel = Not oCust.SetInitials(txtFN)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtFN_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtFN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    Cancel = Not oCust.SetNote(txtNote)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    txtNote = oCust.Note
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtNote_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    txtNote = HandleTextWithBites(txtNote)
    oCust.SetNote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCust.Note
      txtNote.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtNote_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    txtTitle = oCust.Title
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtTitle_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCust.SetTitle (txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oCust.Title
      txtTitle.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtTitle_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetTitle(txtTitle)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


'Private Sub LoadArray()
'    On Error GoTo errHandler
''Dim objItem As d_Customer
'Dim lngIndex As Long
'    XA.ReDim 1, oCust.Addresses.Count, 1, 6
'    For lngIndex = 1 To oCust.Addresses.Count
'        XA.Value(lngIndex, 1) = lngIndex
'        XA.Value(lngIndex, 2) = oCust.Addresses(lngIndex).AddressMailing
'        XA.Value(lngIndex, 3) = CreateRoleString(oCust.Addresses(lngIndex))
'        XA.Value(lngIndex, 4) = oCust.Addresses(lngIndex).GetsCatalogue
'        XA.Value(lngIndex, 5) = oCust.Addresses(lngIndex).Key
'        XA.Value(lngIndex, 6) = oCust.Addresses(lngIndex).ForMailing
'    Next
'    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
'    G1.Array = XA
'    G1.ReBind
'  '  G1.Refresh
'    If XA.UpperBound(1) > 1 Then
'        Me.lblRecords = XA.UpperBound(1) & " addresses"
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.LoadArray"
'End Sub

'Private Sub G1_DblClick()
'    On Error GoTo errHandler
'Dim frm As frmAddress
'Dim lngID As Long
'    Set frm = New frmAddress
'    lngID = val(XA(G1.Bookmark, 5))
'    frm.Component oCust.Addresses.Item(lngID)
'    frm.Show vbModal
'    LoadArray
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.G1_DblClick", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub cmdRemove_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    oCust.Addresses.Remove XA(G1.Bookmark, 5)
'    LoadArray
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.cmdRemove_Click", , EA_NORERAISE
'    HandleError
'End Sub
