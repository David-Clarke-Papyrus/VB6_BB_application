VERSION 5.00
Begin VB.Form frmLoyalty 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Loyalty customer"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   12615
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Customer classification"
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
      Height          =   2145
      Left            =   5430
      TabIndex        =   45
      Top             =   3735
      Width           =   4380
      Begin VB.ListBox lbCC 
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
         Height          =   1230
         Left            =   135
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   795
         Width           =   2700
      End
      Begin VB.ComboBox cboCC 
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
         TabIndex        =   48
         Top             =   375
         Width           =   2745
      End
      Begin VB.CommandButton cmdAddCC 
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
         TabIndex        =   47
         Top             =   345
         Width           =   1305
      End
      Begin VB.CommandButton cmdRemoveCC 
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
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1635
         Width           =   1050
      End
   End
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
      Left            =   2280
      TabIndex        =   17
      Top             =   6540
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
      Height          =   435
      Left            =   2280
      TabIndex        =   16
      Top             =   6240
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
      Left            =   2280
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
      Begin VB.TextBox txtMobile 
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
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2190
         Width           =   2370
      End
      Begin VB.TextBox txtTown 
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
         Top             =   1680
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
         Caption         =   "City/Town"
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
         Left            =   225
         TabIndex        =   36
         Top             =   1740
         Width           =   870
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 3"
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
      Caption         =   "Interest groups"
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
      Height          =   2145
      Left            =   5430
      TabIndex        =   27
      Top             =   1080
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
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1635
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
         Height          =   1230
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
      Height          =   840
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   7065
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
      Height          =   855
      Left            =   7845
      Picture         =   "frmLoyalty.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7020
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8820
      Picture         =   "frmLoyalty.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7005
      Width           =   990
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
      Left            =   7590
      TabIndex        =   3
      Top             =   465
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
      Left            =   105
      TabIndex        =   31
      Top             =   6795
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
      Left            =   5610
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
      Top             =   7020
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
      Left            =   7935
      TabIndex        =   22
      Top             =   210
      Width           =   1065
   End
   Begin VB.Menu mnuACtions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuInactive 
         Caption         =   "&Mark customer as Inactive"
      End
      Begin VB.Menu mnuRemoveFromLoyalty 
         Caption         =   "Remove from Loyalty program"
      End
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

Public Property Get EMail() As String
    EMail = strEMail
End Property


Private Sub chkNotifyLaunches_Click()
    oCust.CustNotifyBookLaunch = (chkNotifyLaunches = 1)
End Sub

Private Sub chkNotifySales_Click()
    oCust.CustNotifyBookSale = (chkNotifySales = 1)
End Sub


Private Sub cboCNTRY_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not cboCNTRY.ListIndex = -1 Then
        oCust.BillTOAddress.CountryID = oCust.BillTOAddress.Countries.Key(cboCNTRY)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.cboCNTRY_Validate(Cancel)", Cancel, EA_NORERAISE
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
    oIG.IGID = oCust.InterestGroupsActive_tl.Key(cboIG)
    oIG.Description = cboIG
    oIG.ApplyEdit
    cboIG.RemoveItem cboIG.ListIndex
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdAddIG_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdAddCC_Click()
    On Error GoTo errHandler
Dim oCC As New a_IG
    If flgLoading Then Exit Sub
    If cboCC = "" Then Exit Sub
    Set oCC = oCust.CustomerTypes.Add
    oCC.BeginEdit
    oCC.TPID = oCust.ID
    oCC.IGID = oCust.CustomerTypes_tl.Key(cboCC)
    oCC.Description = cboCC
    oCC.ApplyEdit
    cboCC.RemoveItem cboCC.ListIndex
    If cboCC.ListCount > 0 Then
        cboCC.ListIndex = 0
    Else
        cboCC.ListIndex = -1
    End If
    LoadTPCCs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdAddCC_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdDuplicates_Click()
    On Error GoTo errHandler
    oCust.LookforDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdDuplicates_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub chkNotifyPromotions_Click()
    oCust.CustNotifyBookPromotion = (chkNotifyPromotions = 1)
End Sub

Private Sub cmdRemoveIG_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbIG = "" Then Exit Sub
    oCust.InterestGroups.Remove oCust.InterestGroups.Key(Me.lbIG)
    cboIG.AddItem Me.lbIG
    cboIG.ListIndex = 0
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdRemoveIG_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdRemoveCC_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbCC = "" Then Exit Sub
    oCust.CustomerTypes.Remove oCust.CustomerTypes.Key(Me.lbCC)
    cboCC.AddItem Me.lbCC
    cboCC.ListIndex = 0
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdRemoveCC_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub Component(pCust As a_Customer)
Dim oAdd As a_Address
    On Error GoTo errHandler
    Set oCust = pCust
    oCust.CustomerTypeID = oPC.Configuration.LoyaltyClubTypeID
    
    Me.Caption = "Customer: " & oCust.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.Component(pCust)", pCust
End Sub
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.EnableOK(pOK)", pOK
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oCust.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
    If flgLoading Then Exit Sub
    'If oCust.IsNew Then
    oCust.LookforDuplicates

    oCust.ApplyEdit lngResult
    If lngResult = 0 Then
        Unload Me
    ElseIf lngResult = 22 Then
        MsgBox "You are trying to save a customer with duplicate values." & vbCrLf & "These are likely to be in the Acc No. field or in the address description fields.", , "Can't save"
    End If
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
    Me.top = 0
    Me.left = 50
    Me.Height = 8200
    Me.Width = 11000
    txtName = oCust.Name
    txtFN = oCust.Initials
    txtAcno = oCust.AcNo
    txtTitle = oCust.Title
    txtNote = oCust.Note
    txtAddressee = oCust.BillTOAddress.Addressee
    Me.txtLine1 = oCust.BillTOAddress.Line1
    Me.txtLine2 = oCust.BillTOAddress.Line2
    Me.txtLine3 = oCust.BillTOAddress.Line3
    Me.txtPCode = oCust.BillTOAddress.pCode
    Me.txtTown = oCust.BillTOAddress.Line6
    Me.txtPhone = oCust.Phone
    Me.txtBusphone = oCust.BillTOAddress.BusPhone
    Me.txtEmail = oCust.BillTOAddress.EMail
    Me.txtMobile = oCust.Mobile
    Me.chkNotifyLaunches = IIf(oCust.CustNotifyBookLaunch, 1, 0)
    Me.chkNotifyPromotions = IIf(oCust.CustNotifyBookPromotion, 1, 0)
    Me.chkNotifySales = IIf(oCust.CustNotifyBookSale, 1, 0)
    LoadCombo cboCNTRY, oCust.BillTOAddress.Countries
    If oCust.BillTOAddress.CountryID > 0 Then
        cboCNTRY.Text = oCust.BillTOAddress.Countries(CStr(oCust.BillTOAddress.CountryID))
    ElseIf oPC.Configuration.LocalCountryID Then
        cboCNTRY.Text = oCust.BillTOAddress.Countries(CStr(oPC.Configuration.LocalCountryID))
        oCust.BillTOAddress.CountryID = oPC.Configuration.LocalCountryID
    End If
    LoadIGs
    LoadCCs
    LoadTPIGs
    LoadTPCCs
    RestrictInterestGroups
    oCust.GetStatus
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadIGs()
    On Error GoTo errHandler
    LoadCombo Me.cboIG, oCust.InterestGroupsActive_tl
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.LoadIGs"
End Sub
Private Sub LoadCCs()
    On Error GoTo errHandler
    LoadCombo Me.cboCC, oCust.CustomerTypes_tl
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.LoadCCs"
End Sub
Private Sub RestrictInterestGroups()
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.LoadTPIGs"
End Sub
Private Sub LoadTPCCs()
    On Error GoTo errHandler
Dim oTPCC As a_IG
    With Me.lbCC
        .Clear
        For Each oTPCC In oCust.CustomerTypes
            .AddItem oTPCC.Description   ', oTPIG.Key
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.LoadTPCCs"
End Sub
'Private Sub lvwAddresses_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub

Private Sub mnuDel_Click()

End Sub

Private Sub oCust_Valid(strMsg As String)
    On Error GoTo errHandler
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.oCust_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_C_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
    HandleError
End Sub


Private Sub ShowDuplicates(pDuplicates As c_C_Customer)
    On Error GoTo errHandler
Dim frm As frmDuplicateCustomers
Dim tmpCust As a_Customer
    
    Set frm = New frmDuplicateCustomers
    frm.Component Me.txtName, pDuplicates
    frm.Show vbModal
    If frm.SelectedCustomer > 0 Then
        Set Forms(0).frmMainCustomerPreview = Nothing
        Set Forms(0).frmMainCustomerPreview = New frmLoyaltyPreview
        Set tmpCust = New a_Customer
        tmpCust.Load frm.SelectedCustomer
        Forms(0).frmMainCustomerPreview.Component tmpCust
    End If
    Unload frm
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.ShowDuplicates(pDuplicates)", pDuplicates
End Sub


Private Sub txtAddressee_GotFocus()
    txtAddressee = oCust.Title & " " & oCust.Initials & " " & oCust.Name
End Sub

Private Sub txtAddressee_Validate(Cancel As Boolean)
On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetAddressee txtAddressee
    If Err Then
      Beep
      intPos = txtLine1.SelStart
      txtAddressee = oCust.BillTOAddress.Addressee
      txtAddressee.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAddressee_Validate", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtLine1_Validate(Cancel As Boolean)
On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetLine1 txtLine1
    If Err Then
      Beep
      intPos = txtLine1.SelStart
      txtLine1 = oCust.BillTOAddress.Line1
      txtLine1.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine1_Validate", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLine2_Validate(Cancel As Boolean)
On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetLine2 txtLine2
    If Err Then
      Beep
      intPos = txtLine2.SelStart
      txtLine2 = oCust.BillTOAddress.Line2
      txtLine2.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine2_Validate", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLine3_Validate(Cancel As Boolean)
On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetLine3 txtLine3
    If Err Then
      Beep
      intPos = txtLine3.SelStart
      txtLine3 = oCust.BillTOAddress.Line3
      txtLine3.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtLine3_Validate", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPCode_Validate(Cancel As Boolean)
On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetPCode txtPCode
    If Err Then
      Beep
      intPos = txtPCode.SelStart
      txtPCode = oCust.BillTOAddress.pCode
      txtPCode.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPCode_Validate", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTown_Validate(Cancel As Boolean)
On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetLine6 txtTown
    If Err Then
      Beep
      intPos = txtTown.SelStart
      txtTown = oCust.BillTOAddress.Line6
      txtTown.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtTown_Validate", , EA_NORERAISE
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

'Private Sub txtMobile_LostFocus()
'    On Error GoTo errHandler
'    txtMobile = oCust.Addresses(1).setMobile
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtMobile_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtMobile_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oCust.setMobile(txtMobile)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtMobile_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtMobile_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    oCust.setMobile (txtMobile)
'    If Err Then
'      Beep
'      intPos = txtMobile.SelStart
'      txtMobile = oCustMobile
'      txtMobile.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyalty.txtMobile_Change", , EA_NORERAISE
'    HandleError
'End Sub
Private Sub txtEMail_LostFocus()
    On Error GoTo errHandler
    txtEmail = oCust.BillTOAddress.EMail
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtEmail_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEMail_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.BillTOAddress.SetEMail(txtEmail)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtEmail_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmail_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetEMail txtEmail
  '  oCust.SetPhone txtEmail    'SetPhone (txtEmail)
    If Err Then
      Beep
      intPos = txtEmail.SelStart
      txtEmail = oCust.BillTOAddress.EMail
      txtEmail.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtEmail_Change", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtBusPhone_LostFocus()
    On Error GoTo errHandler
    txtBusphone = oCust.BillTOAddress.BusPhone
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtBusPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBusPhone_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.BillTOAddress.SetBusPhone(txtBusphone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtBusPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBusPhone_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetBusPhone txtBusphone
  '  oCust.SetPhone txtBusPhone    'SetPhone (txtBusPhone)
    If Err Then
      Beep
      intPos = txtBusphone.SelStart
      txtBusphone = oCust.BillTOAddress.BusPhone
      txtBusphone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtBusPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMobile_LostFocus()
    On Error GoTo errHandler
    txtMobile = oCust.Mobile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtMobile_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMobile_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetMobile(txtMobile)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtMobile_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtMobile_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.SetMobile txtMobile
    If Err Then
      Beep
      intPos = txtMobile.SelStart
      txtMobile = oCust.Phone
      txtMobile.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtMobile_Change", , EA_NORERAISE
    HandleError
End Sub





Private Sub txtPhone_LostFocus()
    On Error GoTo errHandler
    txtPhone = oCust.Phone
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetPhone(txtPhone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.BillTOAddress.SetPhone txtPhone
  '  oCust.SetPhone txtPhone    'SetPhone (txtPhone)
    If Err Then
      Beep
      intPos = txtPhone.SelStart
      txtPhone = oCust.Phone
      txtPhone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
    txtName = oCust.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.SetName (txtName)
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oCust.Name
      txtName.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtName_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetName(txtName)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtName_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtAcno_LostFocus()
    On Error GoTo errHandler
    txtAcno = oCust.AcNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAcno_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.SetAcNO (txtAcno)
    If Err Then
      Beep
      intPos = txtAcno.SelStart
      txtAcno = oCust.AcNo
      txtAcno.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAcno_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetAcNO(txtAcno)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtAcno_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_LostFocus()
    On Error GoTo errHandler
    txtFN = oCust.Initials
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtFN_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.SetInitials (txtFN)
    If Err Then
      Beep
      intPos = txtFN.SelStart
      txtFN = oCust.Initials
      txtFN.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtFN_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetInitials(txtFN)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtFN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.setnote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    txtNote = oCust.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCust.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    txtTitle = oCust.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    oCust.SetTitle (txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oCust.Title
      txtTitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetTitle(txtTitle)
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
