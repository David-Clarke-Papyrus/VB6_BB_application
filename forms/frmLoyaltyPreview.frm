VERSION 5.00
Begin VB.Form frmLoyaltyPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   10095
   Begin VB.ListBox lbCC 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Height          =   750
      Left            =   5730
      TabIndex        =   48
      Top             =   2760
      Width           =   3570
   End
   Begin VB.CheckBox chkNotifyPromotions 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Notify of book promotions"
      Enabled         =   0   'False
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
      Left            =   1410
      TabIndex        =   45
      Top             =   6600
      Width           =   2580
   End
   Begin VB.CheckBox chkNotifyLaunches 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Notify of book launches"
      Enabled         =   0   'False
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
      Left            =   1410
      TabIndex        =   44
      Top             =   6915
      Width           =   2580
   End
   Begin VB.CheckBox chkNotifySales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Notify of book sales"
      Enabled         =   0   'False
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
      Left            =   1410
      TabIndex        =   43
      Top             =   7200
      Width           =   2580
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Addresses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   210
      TabIndex        =   22
      Top             =   1380
      Width           =   4785
      Begin VB.TextBox txtAddressee 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   46
         Top             =   435
         Width           =   3090
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   42
         Top             =   2445
         Width           =   3090
      End
      Begin VB.TextBox txtLine1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   31
         Top             =   840
         Width           =   3090
      End
      Begin VB.TextBox txtLine2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   30
         Top             =   1230
         Width           =   3090
      End
      Begin VB.TextBox txtTown 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   29
         Top             =   1635
         Width           =   3090
      End
      Begin VB.TextBox txtProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   28
         Top             =   2040
         Width           =   3090
      End
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   27
         Top             =   2985
         Width           =   1590
      End
      Begin VB.TextBox txtBusphone 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   26
         Top             =   3810
         Width           =   1920
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   25
         Top             =   3450
         Width           =   1920
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   24
         Top             =   4170
         Width           =   1920
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         TabIndex        =   23
         Top             =   4530
         Width           =   3120
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
         Left            =   105
         TabIndex        =   47
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label9 
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
         TabIndex        =   41
         Top             =   855
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
         TabIndex        =   40
         Top             =   1290
         Width           =   600
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
         TabIndex        =   39
         Top             =   1710
         Width           =   600
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
         Left            =   225
         TabIndex        =   38
         Top             =   2100
         Width           =   870
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
         Left            =   450
         TabIndex        =   37
         Top             =   2475
         Width           =   690
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
         TabIndex        =   36
         Top             =   3015
         Width           =   1095
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
         TabIndex        =   35
         Top             =   3510
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
         TabIndex        =   34
         Top             =   3855
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Mobile"
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
         TabIndex        =   33
         Top             =   4215
         Width           =   1005
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
         TabIndex        =   32
         Top             =   4590
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6825
      Picture         =   "frmLoyaltyPreview.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5910
      Width           =   1050
   End
   Begin VB.CommandButton cmdTPActivity 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Related documents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6690
      Width           =   1800
   End
   Begin VB.PictureBox picNoGO 
      Height          =   420
      Left            =   1245
      Picture         =   "frmLoyaltyPreview.frx":038A
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   19
      Top             =   -120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picOver 
      Height          =   420
      Left            =   1365
      Picture         =   "frmLoyaltyPreview.frx":07CC
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   18
      Top             =   -165
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox PicDrop 
      Height          =   420
      Left            =   675
      Picture         =   "frmLoyaltyPreview.frx":0C0E
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   17
      Top             =   -105
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Height          =   1065
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3960
      Width           =   3585
   End
   Begin VB.ListBox lbIG 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Height          =   750
      Left            =   5745
      TabIndex        =   13
      Top             =   1500
      Width           =   3570
   End
   Begin VB.TextBox txtRecordLastChanged 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Left            =   7710
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   540
      Width           =   1590
   End
   Begin VB.TextBox txtRecordAdded 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Left            =   7710
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   195
      Width           =   1590
   End
   Begin VB.CommandButton cmdShowPurchases 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Purchases"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7065
      Width           =   1785
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   8850
      Picture         =   "frmLoyaltyPreview.frx":1050
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5910
      Width           =   930
   End
   Begin VB.TextBox txtFirstname 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Left            =   3630
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   180
      Width           =   1380
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   180
      Width           =   585
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   7905
      Picture         =   "frmLoyaltyPreview.frx":13DA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5910
      Width           =   930
   End
   Begin VB.TextBox txtAcno 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Left            =   3210
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   690
      Width           =   1785
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Left            =   1005
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2565
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer classification"
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
      Left            =   5760
      TabIndex        =   49
      Top             =   2460
      Width           =   2295
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3540
      X2              =   1065
      Y1              =   -135
      Y2              =   870
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      Left            =   5715
      TabIndex        =   16
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest groups"
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
      Left            =   5775
      TabIndex        =   14
      Top             =   1200
      Width           =   1380
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record last changed: "
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
      Left            =   5880
      TabIndex        =   12
      Top             =   570
      Width           =   1800
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record added: "
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
      Left            =   6360
      TabIndex        =   11
      Top             =   225
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Num."
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
      Left            =   2295
      TabIndex        =   4
      Top             =   765
      Width           =   930
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
      Left            =   345
      TabIndex        =   3
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmLoyaltyPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCust As a_Customer
Dim frmLoy As frmLoyalty
Dim XA As New XArrayDB
Dim vRowBookmark As Variant

Public Sub component(pCust As a_Customer)
    On Error GoTo errHandler
    Set oCust = pCust
    Me.Caption = "Loyalty customer: " & oCust.Name
#If H_CENTRAL <> 1 Then
    Me.cmdShowPurchases.Visible = oPC.Configuration.AntiquarianYN
#End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.Component(pCust)", pCust
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.component(pCust)", pCust
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.cmdClose_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
#If H_CENTRAL <> 1 Then
Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim frm1 As frmTPOldDocs
Dim oDPTP As c_DocsPerTP
Dim lngResult As Long
Dim oSM As z_StockManager

    If MsgBox("You want to delete " & oCust.Fullname, vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Set XA = New XArrayDB
    Set XB = New XArrayDB
    If oCust.OKForDeletion(XA, XB, oDPTP) Then
        If XA.UpperBound(1) > 0 Then
            Set frm1 = New frmTPOldDocs
            frm1.ComponentXA XA, oCust.Fullname, "There are documents belonging to this customer, but they are dated prior to the last stock take and will be deleted if the customer is deleted."
            frm1.Show vbModal
            If Not frm1.ToDelete Then
                Unload frm1
                Exit Sub
            End If
            Unload frm1
        End If
        Set oSM = New z_StockManager
        oSM.DeleteUnusedPTs
        oCust.BeginEdit
        oCust.DeleteCustomer
        oCust.ApplyEdit lngResult
        MsgBox "Customer deleted! Form will close."
        Set oSM = Nothing
        Unload Me
    Else
        MsgBox "There are associated documents which may not be deleted yet. You cannot delete this customer." & vbCrLf & "Use the 'Customer documents button to see details.", , "Can't delete"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.cmdDelete_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub
#End If
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean

    If frmLoy Is Nothing Then
        Set frmLoy = New frmLoyalty
    End If
    blnEdit = True
    oCust.BeginEdit
    frmLoy.component oCust ', lngID
    frmLoy.Show
    
EXIT_Handler:
    Unload Me
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.cmdEdit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdShowPurchases_Click()
    On Error GoTo errHandler
Dim frm As frmCustPurch
Dim oCP As c_SalesPerCustomer

    Set frm = New frmCustPurch
    Set oCP = New c_SalesPerCustomer
    oCP.Load oCust.ID
    frm.component oCP, oCust.Fullname
    frm.Show vbModal
    Set oCP = Nothing
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.cmdShowPurchases_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.cmdShowPurchases_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTPActivity_Click()
    On Error GoTo errHandler
Dim oDPTP As c_DocsPerTP
Dim frm As frmTPActivity

    Set oDPTP = New c_DocsPerTP
    oDPTP.Load oCust.ID
    Set frm = New frmTPActivity
    frm.component oDPTP, oCust.Fullname
    frm.Show vbModal
    
    Unload frm
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.cmdTPActivity_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.cmdTPActivity_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Height = 5100
        Me.Width = 8700
    End If
    txtName = oCust.Name
    txtPhone = oCust.Phone
    txtTitle = oCust.Title
    Me.txtBusphone = oCust.BillTOAddress.BusPhone
    Me.txtEmail = oCust.BillTOAddress.EMail
    Me.txtMobile = oCust.MOBILE
    txtFirstname = oCust.Initials
    txtRecordAdded = oCust.DateRecordAddedF
    txtRecordLastChanged = oCust.DateRecordLastChangedF
    txtAcno = oCust.AcNo
    txtNotes = oCust.Note
    txtAddressee = oCust.BillTOAddress.Addressee
    txtLine1 = oCust.BillTOAddress.Line1
    txtLine2 = oCust.BillTOAddress.Line2
    txtTown = oCust.BillTOAddress.Line5
    txtProvince = oCust.BillTOAddress.Line6
    txtPCode = oCust.BillTOAddress.pCode
    txtPhone = oCust.Phone
    txtCountry = oCust.BillTOAddress.Countries(CStr(oCust.BillTOAddress.CountryID))
    chkNotifyLaunches = IIf(oCust.CustNotifyBookLaunch, 1, 0)
    chkNotifyPromotions = IIf(oCust.CustNotifyBookPromotion, 1, 0)
    chkNotifySales = IIf(oCust.CustNotifyBookSale, 1, 0)
    LoadTPIGs
    LoadTPCCs
    Width = 10300
    Height = 7500
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.Form_Load", , EA_NORERAISE
    HandleError
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
'    ErrorIn "frmLoyaltyPreview.LoadTPIGs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.LoadTPIGs"
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

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLoyaltyPreview.LoadTPCCs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.LoadTPCCs"
End Sub

'Private Sub G1_DblClick()
'Dim frm As frmAddressPreview
'Dim lngID As Long
'    Set frm = New frmAddressPreview
'    lngID = val(XA(G1.Bookmark, 5))
'    frm.Component oCust.Addresses.Item(lngID)
'    frm.Show vbModal
'End Sub
'


'Private Sub G1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'' If the button is up and we get MouseMove, that means
'' we exited the form and tried to drop elsewhere.
'' Reset the drag upon returning.
'    If Button = 0 Then ResetDragDrop
'End Sub
'Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'    If XA(Bookmark, 6) = True Then
'        RowStyle.BackColor = RGB(282, 274, 180)
'    End If
'End Sub
'Private Sub ResetDragDrop()
'' Turn off drag-and-drop by resetting the highlight and data
'' control caption.
'    If G1.MarqueeStyle = dbgSolidCellBorder Then Exit Sub
'    G1.MarqueeStyle = dbgSolidCellBorder
'    G1.MarqueeStyle = dbgSolidCellBorder
''    SB1.SimpleText = "Drag an address"
'End Sub
'Private Sub G1_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
'' Set the current cell to the one being dragged
'    G1.Col = ColIndex
'    G1.Bookmark = RowBookmark
'    vRowBookmark = RowBookmark
'    ' Set up drag operation, such as creating visual effects by
'    ' highlighting the cell or row being dragged.
'            ' Highlight the phone number cell to indicate data
'            ' from the cell is being dragged.
'            G1.MarqueeStyle = dbgHighlightRow
''            SB1.SimpleText = "Dragging an address . . ."
'    ' Use VB manual drag support (put TDBGrid1 into drag mode)
'    G1.Drag vbBeginDrag
'End Sub
'Private Sub G1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'' DragOver provides different visual feedback as we are
'' dragging a row, or just the phone number.
'
'    Dim dragFrom As String
'    Dim overCol As Integer
'    Dim overRow As Long
'
'
'    Select Case State
'        Case vbEnter
'            G1.MarqueeStyle = dbgHighlightRow
'            G1.DragIcon = picOver.Picture
'        Case vbLeave
'            G1.MarqueeStyle = dbgHighlightRow
'            G1.DragIcon = picNoGO.Picture
'        Case vbOver
'            overRow = G1.RowContaining(Y)
'            Debug.Print overRow
'            If overRow >= 0 Then G1.Row = overRow
''            If vRowBookmark = overRow Then
''                G1.DragIcon = picOver.Picture
''            Else
''                G1.DragIcon = PicDrop.Picture
''            End If
'    End Select
'End Sub
'
'Private Sub G1_DragDrop(Source As Control, X As Single, Y As Single)
'    Dim overRow As Long
'        MsgBox "Merging address no: " & vRowBookmark & " Into: " & G1.Bookmark
'End Sub
'
Private Sub Label6_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.Label6_Click", , EA_NORERAISE
    HandleError
End Sub


