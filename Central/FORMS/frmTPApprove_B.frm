VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmTPApprove 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Proposed changes for approval"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   11655
   Begin VB.CheckBox chkOld_LA 
      BackColor       =   &H00D3D3CB&
      Caption         =   "LA"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   10500
      TabIndex        =   88
      Top             =   7800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CheckBox chkOld_SA 
      BackColor       =   &H00D3D3CB&
      Caption         =   "SA"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   9825
      TabIndex        =   87
      Top             =   7800
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CheckBox chkOld_PR 
      BackColor       =   &H00D3D3CB&
      Caption         =   "PR"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   9165
      TabIndex        =   86
      Top             =   7800
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CheckBox chkNew_LA 
      BackColor       =   &H00D3D3CB&
      Caption         =   "LA"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   7860
      TabIndex        =   85
      Top             =   7800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CheckBox chkNew_SA 
      BackColor       =   &H00D3D3CB&
      Caption         =   "SA"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   7185
      TabIndex        =   84
      Top             =   7800
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CheckBox chkNew_PR 
      BackColor       =   &H00D3D3CB&
      Caption         =   "PR"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   6525
      TabIndex        =   83
      Top             =   7800
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   555
      Left            =   1185
      TabIndex        =   80
      Top             =   8055
      Width           =   2670
      Begin VB.OptionButton optNew 
         BackColor       =   &H00D3D3CB&
         Caption         =   "New style"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1425
         TabIndex        =   82
         Top             =   180
         Width           =   1035
      End
      Begin VB.OptionButton optOld 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Old style"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   135
         TabIndex        =   81
         Top             =   165
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.TextBox txtIDNUMN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   4
      Top             =   810
      Width           =   2145
   End
   Begin VB.TextBox txtEmailN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6525
      TabIndex        =   52
      Top             =   7290
      Width           =   2145
   End
   Begin VB.TextBox txtCellN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   49
      Top             =   6885
      Width           =   2145
   End
   Begin VB.TextBox txtFaxN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   46
      Top             =   6480
      Width           =   2145
   End
   Begin VB.TextBox txtNameN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6525
      TabIndex        =   1
      Top             =   405
      Width           =   2145
   End
   Begin VB.TextBox txtInitialsN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   7
      Top             =   1215
      Width           =   2145
   End
   Begin VB.TextBox txtTitleN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   10
      Top             =   1620
      Width           =   2145
   End
   Begin VB.TextBox txtAL1N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   16
      Top             =   2430
      Width           =   2145
   End
   Begin VB.TextBox txtAL2N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   19
      Top             =   2835
      Width           =   2145
   End
   Begin VB.TextBox txtAL3N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   22
      Top             =   3240
      Width           =   2145
   End
   Begin VB.TextBox txtAL4N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   25
      Top             =   3645
      Width           =   2145
   End
   Begin VB.TextBox txtAL5N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   28
      Top             =   4050
      Width           =   2145
   End
   Begin VB.TextBox txtAL6N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   31
      Top             =   4455
      Width           =   2145
   End
   Begin VB.TextBox txtPCodeN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   34
      Top             =   4860
      Width           =   2145
   End
   Begin VB.TextBox txtCountryN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   37
      Top             =   5265
      Width           =   2145
   End
   Begin VB.TextBox txtPhoneN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6555
      TabIndex        =   40
      Top             =   5670
      Width           =   2145
   End
   Begin VB.TextBox txtPhone2N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   43
      Top             =   6075
      Width           =   2145
   End
   Begin VB.TextBox txtAddresseeN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   13
      Top             =   2010
      Width           =   2145
   End
   Begin VB.TextBox txtIDNUMO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "TP_IDNUM"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   6
      Top             =   810
      Width           =   2145
   End
   Begin VB.CommandButton cmdIDNUM 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   795
      Width           =   345
   End
   Begin VB.TextBox txtEMailO 
      Appearance      =   0  'Flat
      DataField       =   "ADD_EMAIL"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   54
      Top             =   7290
      Width           =   2145
   End
   Begin VB.CommandButton cmdEmail 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7275
      Width           =   345
   End
   Begin VB.TextBox txtCellO 
      Appearance      =   0  'Flat
      DataField       =   "TP_CELL"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   51
      Top             =   6870
      Width           =   2145
   End
   Begin VB.CommandButton cmdCell 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6870
      Width           =   345
   End
   Begin VB.TextBox txtFaxO 
      Appearance      =   0  'Flat
      DataField       =   "ADD_FAX"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   48
      Top             =   6465
      Width           =   2145
   End
   Begin VB.CommandButton cmdFax 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6465
      Width           =   345
   End
   Begin VB.TextBox txtDateLastModified 
      Appearance      =   0  'Flat
      DataField       =   "TP_DateLastModified"
      DataSource      =   "Adodc2"
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
      Height          =   375
      Left            =   1950
      TabIndex        =   75
      Text            =   "sdfsdfsdfsdf"
      Top             =   7500
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Remove new record"
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
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8400
      Width           =   2190
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Save"
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8400
      Width           =   2190
   End
   Begin VB.TextBox txtAddresseeO 
      Appearance      =   0  'Flat
      DataField       =   "ADD_ADDRESSEE"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9135
      TabIndex        =   15
      Top             =   2010
      Width           =   2145
   End
   Begin VB.CommandButton cmdAddressee 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2010
      Width           =   345
   End
   Begin VB.ComboBox cboCNTRY 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Height          =   315
      ItemData        =   "frmTPApprove_B.frx":1BB2
      Left            =   9150
      List            =   "frmTPApprove_B.frx":1BB4
      TabIndex        =   39
      Text            =   "cboCNTRY"
      Top             =   5265
      Width           =   2130
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
      Left            =   135
      Picture         =   "frmTPApprove_B.frx":1BB6
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   8040
      Width           =   930
   End
   Begin VB.CommandButton cmdPhone2 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":1C61
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6060
      Width           =   345
   End
   Begin VB.TextBox txtPhone2O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_BUSPHONE"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9165
      TabIndex        =   45
      Top             =   6060
      Width           =   2145
   End
   Begin VB.CommandButton cmdPhone 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":21EB
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5655
      Width           =   345
   End
   Begin VB.TextBox txtPhoneO 
      Appearance      =   0  'Flat
      DataField       =   "ADD_Phone"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   42
      Top             =   5655
      Width           =   2145
   End
   Begin VB.CommandButton cmdCountry 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":2775
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5250
      Width           =   345
   End
   Begin VB.CommandButton cmdPCode 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":2CFF
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4845
      Width           =   345
   End
   Begin VB.TextBox txtPCodeO 
      Appearance      =   0  'Flat
      DataField       =   "ADD_PCode"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   36
      Top             =   4860
      Width           =   2145
   End
   Begin VB.CommandButton cmdAL6 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":3289
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4440
      Width           =   345
   End
   Begin VB.TextBox txtAL6O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L6"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   33
      Top             =   4455
      Width           =   2145
   End
   Begin VB.CommandButton cmdAL5 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":3813
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4035
      Width           =   345
   End
   Begin VB.TextBox txtAL5O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L5"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   30
      Top             =   4065
      Width           =   2145
   End
   Begin VB.CommandButton cmdAL4 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":3D9D
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3630
      Width           =   345
   End
   Begin VB.TextBox txtAL4O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L4"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   27
      Top             =   3645
      Width           =   2145
   End
   Begin VB.CommandButton cmdAL3 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":4327
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3225
      Width           =   345
   End
   Begin VB.TextBox txtAL3O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L3"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9120
      TabIndex        =   24
      Top             =   3240
      Width           =   2145
   End
   Begin VB.CommandButton cmdAL2 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":48B1
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2820
      Width           =   345
   End
   Begin VB.TextBox txtAL2O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L2"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   21
      Top             =   2835
      Width           =   2145
   End
   Begin VB.CommandButton cmdAL1 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":4E3B
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2415
      Width           =   345
   End
   Begin VB.TextBox txtAL1O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L1"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   18
      Top             =   2430
      Width           =   2145
   End
   Begin VB.CommandButton cmdTitle 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":53C5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1605
      Width           =   345
   End
   Begin VB.TextBox txtTitleO 
      Appearance      =   0  'Flat
      DataField       =   "TP_TITLE"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   12
      Top             =   1620
      Width           =   2145
   End
   Begin VB.CommandButton cmdInitials 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":594F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   345
   End
   Begin VB.TextBox txtInitialsO 
      Appearance      =   0  'Flat
      DataField       =   "TP_INITIALS"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   9
      Top             =   1215
      Width           =   2145
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00D8CEAB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8745
      Picture         =   "frmTPApprove_B.frx":5ED9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   390
      Width           =   345
   End
   Begin VB.TextBox txtNameO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "TP_NAME"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   3
      Top             =   405
      Width           =   2145
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   6270
      Left            =   90
      OleObjectBlob   =   "frmTPApprove_B.frx":6463
      TabIndex        =   0
      Top             =   375
      Width           =   5085
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "I.D.Number"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   79
      Top             =   812
      Width           =   1335
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   5115
      TabIndex        =   78
      Top             =   7335
      Width           =   1335
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   5070
      TabIndex        =   77
      Top             =   6915
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Height          =   315
      Left            =   5085
      TabIndex        =   76
      Top             =   6510
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   5115
      TabIndex        =   74
      Top             =   2033
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Existing data"
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
      Height          =   315
      Left            =   9165
      TabIndex        =   73
      Top             =   105
      Width           =   2040
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "New data"
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
      Height          =   285
      Left            =   6525
      TabIndex        =   72
      Top             =   75
      Width           =   2040
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (bus)"
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
      Height          =   315
      Left            =   5085
      TabIndex        =   71
      Top             =   6103
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   5100
      TabIndex        =   70
      Top             =   5696
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   5100
      TabIndex        =   69
      Top             =   5289
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Post code"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   68
      Top             =   4882
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   5100
      TabIndex        =   67
      Top             =   4475
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Suburb/Town"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   66
      Top             =   4068
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "A.L.4"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   65
      Top             =   3661
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "A.L.3"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   64
      Top             =   3254
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "A.L.2"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   63
      Top             =   2847
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "A.L.1"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   62
      Top             =   2440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   5100
      TabIndex        =   61
      Top             =   1626
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Initials"
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
      Height          =   315
      Left            =   5100
      TabIndex        =   60
      Top             =   1219
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   5100
      TabIndex        =   59
      Top             =   405
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "New data (summary)"
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
      Height          =   315
      Left            =   105
      TabIndex        =   58
      Top             =   120
      Width           =   2790
   End
End
Attribute VB_Name = "frmTPApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsNew As ADODB.Recordset
Dim rsOld As ADODB.Recordset
Dim rsIG As ADODB.Recordset

Dim flgLoaded As Boolean
Dim tlCTRY As z_TextList
Dim flgLoading As Boolean
Dim bmForApproval
Dim X1 As New XArrayDB
Dim X2 As New XArrayDB
Dim strSource As String



Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    LoadCombo cboCNTRY, tlCTRY
    Me.Width = 11600
    Me.Height = 9700
    Me.top = 160
    Me.left = 200
    
    If MsgBox("Import from Papyrus stores?", vbYesNo, "Source") = vbYes Then
        strSource = "PAPYRUS"
    Else
        strSource = "WORDSTOCK"
    End If
    
    LoadX1 strSource
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadX1(pSource As String)
Dim lngIndex As Long
Dim oSQL As New z_SQL
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set rsNew = Nothing
    Set rsNew = New ADODB.Recordset
    rsNew.CursorLocation = adUseClient
    oPC.COShort.Execute "DELETE FROM tADD_ForApproval2 FROM tADD_ForApproval2 LEFT JOIN tTP_ForApproval2 ON TP_ID = ADD_TP_ID WHERE TP_ID IS NULL"
    oPC.COShort.Execute "DELETE FROM tADD_ForApproval FROM tADD_ForApproval LEFT JOIN tTP_ForApproval ON TP_ID = ADD_TP_ID WHERE TP_ID IS NULL"
    If pSource = "PAPYRUS" Then
       ' oSQL.GetDynamicRecordset "Select * FROM tTP_Forapproval2 JOIN tADD_ForApproval2 ON ADD_TP_ID = TP_ID ORDER BY TP_NAME", enText, Array(), "", rsNew
        oSQL.GetDynamicRecordset "Select * FROM vForApproval2_Std ORDER BY TP_NAME", enText, "", rsNew
    Else
       ' oSQL.GetDynamicRecordset "Select * FROM tTP_Forapproval JOIN tADD_ForApproval ON ADD_TP_ID = TP_ID ORDER BY TP_NAME", enText, Array(), "", rsNew
        oSQL.GetDynamicRecordset "Select * FROM vForApproval_Std ORDER BY TP_NAME", enText, "", rsNew
    End If
    X1.ReDim 1, rsNew.RecordCount, 1, 5
    lngIndex = 1
    Do While Not rsNew.EOF
        X1(lngIndex, 1) = FNS(rsNew.Fields("TP_ACNO"))
        X1(lngIndex, 2) = FNS(rsNew.Fields("TP_Name"))
        X1(lngIndex, 3) = FNS(rsNew.Fields("TP_INITIALS"))
        X1(lngIndex, 4) = FNS(rsNew.Fields("TP_TITLE"))
     '   X1(lngIndex, 5) = FNS(CStr(rsNew.Fields(0)))
        lngIndex = lngIndex + 1
        rsNew.MoveNext
    Loop
    
    X1.QuickSort 1, X1.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set G.Array = X1
    G.ReBind
   DoEvents
    If X1.UpperBound(1) > 0 Then
        rsNew.Find "TP_ACNO = '" & FNS(X1(G.Bookmark, 1)) & "'", , adSearchForward, 1
        LoadNewData
    End If

End Sub
Private Sub cmdRemove_Click()
    On Error GoTo errHandler
Dim bkmark
    ClearBackgroundColour
    If X1.UpperBound(1) = 0 Then Exit Sub
  '  rsNew.Delete
    If strSource = "PAPYRUS" Then
        oPC.COShort.Execute "DELETE FROM tTP_ForApproval2 WHERE TP_ACNO = '" & rsNew.Fields("TP_ACNO") & "'"
    Else
        oPC.COShort.Execute "DELETE FROM tTP_ForApproval WHERE TP_ACNO = '" & rsNew.Fields("TP_ACNO") & "'"
    End If
    G.Delete
    If Not IsNull(G.Bookmark) Then
        rsNew.Find "TP_ACNO = '" & FNS(X1(G.Bookmark, 1)) & "'", , adSearchForward, 1
        If Not (rsNew.EOF Or rsNew.BOF) Then
            LoadNewData
        End If
    End If
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
Dim oSQL As New z_SQL

    On Error GoTo errHandler
    If (FNN(rsIG.Fields("PR")) = 1 Or FNN(rsIG.Fields("SA")) = 1 Or FNN(rsIG.Fields("LA")) = 1 Or FNN(rsIG.Fields("NL")) = 1) And (FNS(rsOld.Fields("ADD_EMAIL")) = "") Then
        If MsgBox("Warning: The email address is missing and there are interest groups selected", vbInformation + vbOKCancel, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    rsOld.Update
    If FNS(rsOld.Fields("ADD_EMAIL")) > "" Then
        oSQL.RunProc "SetNewsletterTrue", Array(FNN(rsOld.Fields("TP_ID")), 1), "", oPC.COShort
    Else
        oSQL.RunProc "SetNewsletterTrue", Array(FNN(rsOld.Fields("TP_ID")), 0), "", oPC.COShort
    End If

    txtDateLastModified = Format(Now(), "dd/mm/yyyy HH:nn")
    ClearBackgroundColour
    If X1.UpperBound(1) = 0 Then Exit Sub
    If strSource = "PAPYRUS" Then
        oPC.COShort.Execute "DELETE FROM tTP_ForApproval2 WHERE TP_ACNO = '" & rsNew.Fields("TP_ACNO") & "'"
    Else
        oPC.COShort.Execute "DELETE FROM tTP_ForApproval WHERE TP_ACNO = '" & rsNew.Fields("TP_ACNO") & "'"
    End If
  '  rsNew.Delete
    G.Delete
    
    rsNew.Find "TP_ACNO = '" & FNS(X1(G.Bookmark, 1)) & "'", , adSearchForward, 1
    
    If Not (rsNew.EOF Or rsNew.BOF) Then
        LoadNewData
    End If
    
    Exit Sub
errHandler:
    ErrPreserve
    If Err = -2147217887 Then
        MsgBox "Cannot save"
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If X1.UpperBound(1) = 0 Then Exit Sub
    rsNew.Find "TP_ACNO = '" & FNS(X1(G.Bookmark, 1)) & "'", , adSearchForward, 1
    LoadNewData
End Sub

Sub LoadNewData()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    
    flgLoading = True
    
'    If rsOld.State = 1 Then rsOld.Close
    Set rsOld = Nothing
    Set rsOld = New ADODB.Recordset
    rsOld.CursorLocation = adUseClient
    oSQL.GetDynamicRecordset "Select * FROM tTP JOIN tADD ON ADD_TP_ID = TP_ID WHERE TP_ACNO = '" & FNS(rsNew.Fields("TP_ACNO")) & "'", enText, "", rsOld
    
'    If rsIG.State = 1 Then rsIG.Close
    Set rsIG = Nothing
    Set rsIG = New ADODB.Recordset
    rsIG.CursorLocation = adUseClient
    oSQL.GetDynamicRecordset "Select * FROM vInterestGroupMembership WHERE TP_ACNO = '" & FNS(rsNew.Fields("TP_ACNO")) & "'", enText, "", rsIG
    If Not rsOld.EOF Then
        Me.txtNameO = FNS(rsOld.Fields("TP_NAME"))
        Me.txtAddresseeO = FNS(rsOld.Fields("ADD_Addressee"))
        Me.txtAL1O = FNS(rsOld.Fields("ADD_L1"))
        Me.txtAL2O = FNS(rsOld.Fields("ADD_L2"))
        Me.txtAL3O = FNS(rsOld.Fields("ADD_L3"))
        Me.txtAL4O = FNS(rsOld.Fields("ADD_L4"))
        Me.txtAL5O = FNS(rsOld.Fields("ADD_L5"))
        Me.txtAL6O = FNS(rsOld.Fields("ADD_L6"))
        Me.txtCellO = FNS(rsOld.Fields("TP_CELL"))
        Me.txtEMailO = FNS(rsOld.Fields("ADD_EMAIL"))
        Me.txtFaxO = FNS(rsOld.Fields("ADD_FAX"))
        Me.txtIDNUMO = FNS(rsOld.Fields("TP_IDNUM"))
        Me.txtInitialsO = FNS(rsOld.Fields("TP_INITIALS"))
        Me.txtPCodeO = FNS(rsOld.Fields("ADD_PCode"))
        Me.txtPhone2O = FNS(rsOld.Fields("ADD_BUSPHONE"))
        Me.txtPhoneO = FNS(rsOld.Fields("ADD_Phone"))
        Me.txtTitleO = FNS(rsOld.Fields("TP_TITLE"))
    End If
      flgLoading = False

    
    txtNameN = FNS(rsNew.Fields("TP_Name"))
    If txtNameO <> txtNameN Then
        txtNameN.ForeColor = vbRed
    Else
        txtNameN.ForeColor = txtNameO.ForeColor
    End If

    txtInitialsN = FNS(rsNew.Fields("TP_Initials"))
    If txtInitialsO <> txtInitialsN Then
        txtInitialsN.ForeColor = vbRed
    Else
        txtInitialsN.ForeColor = txtInitialsO.ForeColor
    End If
    
    txtTitleN = FNS(rsNew.Fields("TP_Title"))
    If txtTitleO <> txtTitleN Then
        txtTitleN.ForeColor = vbRed
    Else
        txtTitleN.ForeColor = txtTitleO.ForeColor
    End If
'MsgBox "POS 3"
    txtAddresseeN = FNS(rsNew.Fields("ADD_Addressee"))
    If txtAddresseeO <> txtAddresseeN Then
        txtAddresseeN.ForeColor = vbRed
    Else
        txtAddresseeN.ForeColor = txtAddresseeO.ForeColor
    End If
    
    txtAL1N = FNS(rsNew.Fields("ADD_L1"))
    If txtAL1O <> txtAL1N Then
        txtAL1N.ForeColor = vbRed
    Else
        txtAL1N.ForeColor = txtAL1O.ForeColor
    End If
    
    txtAL2N = FNS(rsNew.Fields("ADD_L2"))
    If txtAL2O <> txtAL2N Then
        txtAL2N.ForeColor = vbRed
    Else
        txtAL2N.ForeColor = txtAL2O.ForeColor
    End If
    
    txtAL3N = FNS(rsNew.Fields("ADD_L3"))
    If txtAL3O <> txtAL3N Then
        txtAL3N.ForeColor = vbRed
    Else
        txtAL3N.ForeColor = txtAL3O.ForeColor
    End If
    
    txtAL4N = FNS(rsNew.Fields("ADD_L4"))
    If txtAL4O <> txtAL4N Then
        txtAL4N.ForeColor = vbRed
    Else
        txtAL4N.ForeColor = txtAL4O.ForeColor
    End If
    
    txtAL5N = FNS(rsNew.Fields("ADD_L5"))
    If txtAL5O <> txtAL5N Then
        txtAL5N.ForeColor = vbRed
    Else
        txtAL5N.ForeColor = txtAL5O.ForeColor
    End If
    
    txtAL6N = FNS(rsNew.Fields("ADD_L6"))
    If txtAL6O <> txtAL6N Then
        txtAL6N.ForeColor = vbRed
    Else
        txtAL6N.ForeColor = txtAL6O.ForeColor
    End If
    
    txtPCodeN = FNS(rsNew.Fields("ADD_PCode"))
    If txtPCodeO <> txtPCodeN Then
        txtPCodeN.ForeColor = vbRed
    Else
        txtPCodeN.ForeColor = txtPCodeO.ForeColor
    End If
    If Not rsOld.EOF Then
        Me.cboCNTRY.Text = FNS(tlCTRY.Item(FNN(rsOld.Fields("ADD_CNTRY_ID"))))
    End If
    If Not rsOld.EOF Then
        txtCountryN = tlCTRY.Item(FNN(rsOld.Fields("ADD_CNTRY_ID")))
    End If
    If cboCNTRY <> txtCountryN Then
        txtCountryN.ForeColor = vbRed
    Else
        txtCountryN.ForeColor = cboCNTRY.ForeColor
    End If
    
    If PhoneFormat(txtPhoneN, "") = PhoneFormat(FNS(rsNew.Fields("ADD_Phone")), "") Then
   ' If txtPhoneO <> txtPhoneN Then
        txtPhoneN.ForeColor = vbRed
    Else
        txtPhoneN.ForeColor = txtPhoneO.ForeColor
    End If
    
    txtCellN = FNS(rsNew.Fields("TP_Cell"))
    If txtCellO <> txtCellN Then
        txtCellN.ForeColor = vbRed
    Else
        txtCellN.ForeColor = txtCellO.ForeColor
    End If
    
    txtEmailN = FNS(rsNew.Fields("ADD_EmAIL"))
    If txtEMailO <> txtEmailN Then
        txtEmailN.ForeColor = vbRed
    Else
        txtEmailN.ForeColor = txtEMailO.ForeColor
    End If
    
    txtFaxN = FNS(rsNew.Fields("ADD_FAX"))
    If txtFaxO <> txtFaxN Then
        txtFaxN.ForeColor = vbRed
    Else
        txtFaxN.ForeColor = txtFaxO.ForeColor
    End If
    
    txtIDNUMN = FNS(rsNew.Fields("TP_IDNUM"))
    If txtIDNUMO <> txtIDNUMN Then
        txtIDNUMN.ForeColor = vbRed
    Else
        txtIDNUMN.ForeColor = txtIDNUMO.ForeColor
    End If
    
    txtPhone2N = FNS(rsNew.Fields("ADD_BUSPHONE"))
    
'        Me.chkNew_PR = IIf(rsNew.Fields("PR"), 1, 0)
'        Me.chkNew_SA = IIf(rsNew.Fields("SA"), 1, 0)
'        Me.chkNew_LA = IIf(rsNew.Fields("LA"), 1, 0)
'
'        Me.chkOld_PR = IIf(rsIG.Fields("PR"), 1, 0)
'        Me.chkOld_SA = IIf(rsIG.Fields("SA"), 1, 0)
'        Me.chkOld_LA = IIf(rsIG.Fields("LA"), 1, 0)
'
    
    If rsOld.EOF Then
        flgLoading = True
        txtNameO = ""
        Me.txtIDNUMO = ""
        txtInitialsO = ""
        txtTitleO = ""
        txtAddresseeO = ""
        txtAL1O = ""
        txtAL2O = ""
        txtAL3O = ""
        txtAL4O = ""
        txtAL5O = ""
        txtAL6O = ""
        txtPCodeO = ""
        'txtAL6O = ""
        txtPhoneO = ""
        txtCellO = ""
        txtEmailN = ""
        txtFaxN = ""
        txtIDNUMN = ""
        
        txtPhone2N = ""
        flgLoading = False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmTPApprove.LoadNewData(rsNew)", rsNew
'    Exit Sub
'    Resume
    Exit Sub
errHandler:
    ErrorIn "frmTPApprove.LoadNewData"
End Sub


Private Sub cboCNTRY_Click()
    On Error GoTo errHandler
'    tlCTRY.Item (FNN(rsOLD.Fields("ADD_CNTRY_ID")))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cboCNTRY_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddressee_Click()
    On Error GoTo errHandler
    txtAddresseeO = txtAddresseeN
    rsOld.Fields("ADD_Addressee") = txtAddresseeN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAddressee_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL1_Click()
    On Error GoTo errHandler
    txtAL1O = txtAL1N
    rsOld.Fields("ADD_L1") = txtAL1N
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL2_Click()
    On Error GoTo errHandler
    txtAL2O = txtAL2N
    rsOld.Fields("ADD_L2") = txtAL2N
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL3_Click()
    On Error GoTo errHandler
    txtAL3O = txtAL3N

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL3_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL4_Click()
    On Error GoTo errHandler
    txtAL4O = txtAL4N
    rsOld.Fields("ADD_L4") = txtAL4N
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL4_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL5_Click()
    On Error GoTo errHandler
    txtAL5O = txtAL5N
    rsOld.Fields("ADD_L5") = txtAL5N

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL5_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL6_Click()
    On Error GoTo errHandler
    txtAL6O = txtAL6N
    rsOld.Fields("ADD_L6") = txtAL6N

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL6_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdIDNum_Click()
    On Error GoTo errHandler
    txtIDNUMO = txtIDNUMN
    rsOld.Fields("TP_IDNUM") = txtIDNUMN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdIDNum_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCell_Click()
    On Error GoTo errHandler
    txtCellO = txtCellN
    rsOld.Fields("TP_CELL") = txtCellN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdCell_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPhone2_Click()
    On Error GoTo errHandler
    txtPhone2O = txtPhone2N
    rsOld.Fields("ADD_BUSPHONE") = txtPhone2N
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdPhone2_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdFax_Click()
    On Error GoTo errHandler
    txtFaxO = txtFaxN
    rsOld.Fields("ADD_FAX") = txtFaxN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdFax_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEmail_Click()
    On Error GoTo errHandler
    txtEMailO = txtEmailN
    rsOld.Fields("ADD_EMAIL") = txtEmailN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdEmail_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdInitials_Click()
    On Error GoTo errHandler
    txtInitialsO.Text = txtInitialsN
    rsOld.Fields("TP_INITIALS") = txtInitialsN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdInitials_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdName_Click()
    On Error GoTo errHandler
    txtNameO = txtNameN
    rsOld.Fields("TP_NAME") = txtNameN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdName_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPCode_Click()
    On Error GoTo errHandler
    txtPCodeO = txtPCodeN
    rsOld.Fields("ADD_PCODE") = txtPCodeN

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdPCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPhone_Click()
    On Error GoTo errHandler
    txtPhoneO = txtPhoneN
    rsOld.Fields("ADD_PHONE") = txtPhoneN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdPhone_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountry_Click()
Dim strCountryCode As String
Dim ky As Variant

    strCountryCode = txtPhoneN
    ky = tlCTRY.Key(strCountryCode)
    If ky > "" Then
        rsOld.Fields("ADD_CNTRY_ID") = ky
        Me.cboCNTRY.Text = strCountryCode
    End If
End Sub
Private Sub cmdTitle_Click()
    On Error GoTo errHandler
    txtTitleO = txtTitleN
    rsOld.Fields("TP_TITLE") = txtTitleN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdTitle_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    flgLoaded = False
    Set tlCTRY = New z_TextList
    tlCTRY.Load ltCountry
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetDefaultWidths()
    On Error GoTo errHandler
Dim i As Integer
'    For i = 1 To G.Columns.Count
'        G.Columns(i - 1).Width = 500
'    Next
'    For i = 1 To GG.Columns.Count
'        GG.Columns(i - 1).Width = 500
'    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.SetDefaultWidths"
End Sub


Private Sub ClearBackgroundColour()
    On Error GoTo errHandler
    txtAddresseeO.BackColor = vbWhite
    txtAL1O.BackColor = vbWhite
    txtAL2O.BackColor = vbWhite
    txtAL3O.BackColor = vbWhite
    txtAL4O.BackColor = vbWhite
    txtAL5O.BackColor = vbWhite
    txtAL6O.BackColor = vbWhite
    txtCellO.BackColor = vbWhite
    txtInitialsO.BackColor = vbWhite
    txtNameO.BackColor = vbWhite
    txtPhoneO.BackColor = vbWhite
    txtPCodeO.BackColor = vbWhite
    txtTitleO.BackColor = vbWhite
    txtEMailO.BackColor = vbWhite
    txtFaxO.BackColor = vbWhite
    txtIDNUMO.BackColor = vbWhite
    txtPhone2O.BackColor = vbWhite

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.ClearBackgroundColour"
End Sub





Private Sub txtAddresseeO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_Addressee") = txtAddresseeO
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAddresseeO.BackColor = IIf(txtAddresseeO = FNS(rsOld.Fields("ADD_Addressee").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAddresseeO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL1O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_L1") = txtAL1O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAL1O.BackColor = IIf(txtAL1O = FNS(rsOld.Fields("ADD_L1").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL1O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL2O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_L2") = txtAL2O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAL2O.BackColor = IIf(txtAL2O = FNS(rsOld.Fields("ADD_L2").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL2O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL3O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_L3") = txtAL3O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAL3O.BackColor = IIf(txtAL3O = FNS(rsOld.Fields("ADD_L3").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL3O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL4O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_L4") = txtAL4O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAL4O.BackColor = IIf(txtAL4O = FNS(rsOld.Fields("ADD_L4").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL4O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL5O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_L5") = txtAL5O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAL5O.BackColor = IIf(txtAL5O = FNS(rsOld.Fields("ADD_L5").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL5O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL6O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("ADD_L6") = txtAL6O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtAL6O.BackColor = IIf(txtAL6O = FNS(rsOld.Fields("ADD_L6").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL6O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCellO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("TP_CELL") = txtCellO
    If Not (rsOld.EOF Or rsOld.BOF) Then
        txtCellO.BackColor = IIf(txtCellO = FNS(rsOld.Fields("TP_CELL").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtCellO_Change", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtInitialsO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    rsOld.Fields("TP_INITIALS") = txtInitialsO
    If Not (rsOld.EOF Or rsOld.BOF) Then
        txtInitialsO.BackColor = IIf(txtInitialsO = FNS(rsOld.Fields("TP_INITIALS").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtInitialsO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtNameO_Change()
    On Error GoTo errHandler
    If flgLoading Then
        Exit Sub
    End If
    rsOld.Fields("TP_Name") = txtNameO
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtNameO.BackColor = IIf(txtNameO = FNS(rsOld.Fields("TP_Name").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtNameO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPCodeO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("ADD_PCODE") = txtPCodeO
   If Not (rsOld.EOF Or rsOld.BOF) Then
        txtPCodeO.BackColor = IIf(txtPCodeO = FNS(rsOld.Fields("ADD_PCODE").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtPCodeO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhoneO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("ADD_PHONE") = txtPhoneO
    If Not (rsOld.EOF Or rsOld.BOF) Then
        txtPhoneO.BackColor = IIf(txtPhoneO = FNS(rsOld.Fields("ADD_PHONE").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtPhoneO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitleO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("TP_TITLE") = txtTitleO
    If Not (rsOld.EOF Or rsOld.BOF) Then
        txtTitleO.BackColor = IIf(txtTitleO = FNS(rsOld.Fields("TP_TITLE").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtTitleO_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtPhone2O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("ADD_BUSPHONE") = txtPhone2O
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtPhone2O.BackColor = IIf(txtPhone2O = FNS(rsOld.Fields("ADD_BUSPHONE").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtPhone2O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtIDNUMO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("TP_IDNUM") = txtIDNUMO
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtIDNUMO.BackColor = IIf(txtIDNUMO = FNS(rsOld.Fields("TP_IDNUM").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtIDNUMO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFaxO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("ADD_FAX") = txtFaxO
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtFaxO.BackColor = IIf(txtFaxO = FNS(rsOld.Fields("ADD_FAX").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtFaxO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEMailO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
     rsOld.Fields("ADD_EMAIL") = txtEMailO
    If Not (rsOld.EOF Or rsOld.BOF) Then
            txtEMailO.BackColor = IIf(txtEMailO = FNS(rsOld.Fields("ADD_EMAIL").OriginalValue), vbWhite, &HC0C0FF)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtEMailO_Change", , EA_NORERAISE
    HandleError
End Sub

