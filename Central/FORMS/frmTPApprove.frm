VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTPApprove 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Proposed changes for approval"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   11655
   Begin VB.TextBox txtIDNUMN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   79
      Top             =   810
      Width           =   2145
   End
   Begin VB.TextBox txtEmailN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   78
      Top             =   7290
      Width           =   2145
   End
   Begin VB.TextBox txtCellN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   77
      Top             =   6885
      Width           =   2145
   End
   Begin VB.TextBox txtFaxN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   76
      Top             =   6480
      Width           =   2145
   End
   Begin VB.TextBox txtNameN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6525
      TabIndex        =   75
      Top             =   405
      Width           =   2145
   End
   Begin VB.TextBox txtInitialsN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   74
      Top             =   1215
      Width           =   2145
   End
   Begin VB.TextBox txtTitleN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   73
      Top             =   1620
      Width           =   2145
   End
   Begin VB.TextBox txtAL1N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   72
      Top             =   2430
      Width           =   2145
   End
   Begin VB.TextBox txtAL2N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   71
      Top             =   2835
      Width           =   2145
   End
   Begin VB.TextBox txtAL3N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   70
      Top             =   3240
      Width           =   2145
   End
   Begin VB.TextBox txtAL4N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   69
      Top             =   3645
      Width           =   2145
   End
   Begin VB.TextBox txtAL5N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   68
      Top             =   4050
      Width           =   2145
   End
   Begin VB.TextBox txtAL6N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   67
      Top             =   4455
      Width           =   2145
   End
   Begin VB.TextBox txtPCodeN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   66
      Top             =   4860
      Width           =   2145
   End
   Begin VB.TextBox txtCountryN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   65
      Top             =   5265
      Width           =   2145
   End
   Begin VB.TextBox txtPhoneN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   64
      Top             =   5670
      Width           =   2145
   End
   Begin VB.TextBox txtPhone2N 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   63
      Top             =   6075
      Width           =   2145
   End
   Begin VB.TextBox txtAddresseeN 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6540
      TabIndex        =   62
      Top             =   2025
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
      TabIndex        =   60
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
      Picture         =   "frmTPApprove.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
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
      TabIndex        =   57
      Top             =   7275
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
      Picture         =   "frmTPApprove.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   56
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
      TabIndex        =   54
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
      Picture         =   "frmTPApprove.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   53
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
      TabIndex        =   51
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
      Picture         =   "frmTPApprove.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   8070
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
      TabIndex        =   47
      Top             =   8070
      Width           =   2190
   End
   Begin VB.TextBox txtAddresseeO 
      Appearance      =   0  'Flat
      DataField       =   "ADD_ADDRESSEE"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   45
      Top             =   2025
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
      Picture         =   "frmTPApprove.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2010
      Width           =   345
   End
   Begin VB.ComboBox cboCNTRY 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Height          =   315
      ItemData        =   "frmTPApprove.frx":1BB2
      Left            =   9150
      List            =   "frmTPApprove.frx":1BB4
      TabIndex        =   43
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
      Picture         =   "frmTPApprove.frx":1BB6
      Style           =   1  'Graphical
      TabIndex        =   40
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
      Picture         =   "frmTPApprove.frx":1C61
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6060
      Width           =   345
   End
   Begin VB.TextBox txtPhone2O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_BUSPHONE"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   37
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
      Picture         =   "frmTPApprove.frx":21EB
      Style           =   1  'Graphical
      TabIndex        =   32
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
      TabIndex        =   31
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
      Picture         =   "frmTPApprove.frx":2775
      Style           =   1  'Graphical
      TabIndex        =   30
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
      Picture         =   "frmTPApprove.frx":2CFF
      Style           =   1  'Graphical
      TabIndex        =   29
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
      TabIndex        =   28
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
      Picture         =   "frmTPApprove.frx":3289
      Style           =   1  'Graphical
      TabIndex        =   27
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
      TabIndex        =   26
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
      Picture         =   "frmTPApprove.frx":3813
      Style           =   1  'Graphical
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   4050
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
      Picture         =   "frmTPApprove.frx":3D9D
      Style           =   1  'Graphical
      TabIndex        =   21
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
      TabIndex        =   20
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
      Picture         =   "frmTPApprove.frx":4327
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3225
      Width           =   345
   End
   Begin VB.TextBox txtAL3O 
      Appearance      =   0  'Flat
      DataField       =   "ADD_L3"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9150
      TabIndex        =   16
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
      Picture         =   "frmTPApprove.frx":48B1
      Style           =   1  'Graphical
      TabIndex        =   15
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
      TabIndex        =   14
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
      Picture         =   "frmTPApprove.frx":4E3B
      Style           =   1  'Graphical
      TabIndex        =   11
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
      TabIndex        =   10
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
      Picture         =   "frmTPApprove.frx":53C5
      Style           =   1  'Graphical
      TabIndex        =   9
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
      TabIndex        =   8
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
      Picture         =   "frmTPApprove.frx":594F
      Style           =   1  'Graphical
      TabIndex        =   5
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
      TabIndex        =   4
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
      Picture         =   "frmTPApprove.frx":5ED9
      Style           =   1  'Graphical
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   405
      Width           =   2145
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   3555
      Top             =   6735
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmTPApprove.frx":6463
      Height          =   6270
      Left            =   90
      OleObjectBlob   =   "frmTPApprove.frx":6478
      TabIndex        =   1
      Top             =   375
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   9150
      Top             =   7635
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483635
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      TabIndex        =   61
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
      TabIndex        =   58
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
      Left            =   5085
      TabIndex        =   55
      Top             =   6917
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
      TabIndex        =   52
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
      TabIndex        =   46
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
      TabIndex        =   42
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
      TabIndex        =   41
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
      TabIndex        =   39
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   4882
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "A.L.6"
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
      TabIndex        =   33
      Top             =   4475
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "A.L.5"
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   0
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
Dim rs As ADODB.Recordset
Dim flgLoaded As Boolean
Dim tlCTRY As z_TextList
Dim flgLoading As Boolean

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    txtDateLastModified = Format(Now(), "dd/mm/yyyy HH:nn")
    Adodc2.Recordset.Update
    ClearBackgroundColour
    Me.Adodc1.Recordset.Delete
    Me.Adodc1.Recordset.Update
    G.Refresh
    Me.Adodc1.Refresh
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

Private Sub Adodc2_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo errHandler
    If Not Adodc2.Recordset.EOF Then
        Adodc2.Recordset.CancelUpdate
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.Adodc2_WillMove(adReason,adStatus,pRecordset)", Array(adReason, adStatus, _
         pRecordset), EA_NORERAISE
    HandleError
End Sub

Private Sub Adodc2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo errHandler
    If Not Adodc2.Recordset.EOF Then
        Adodc2.Caption = CStr(pRecordset.AbsolutePosition) & " of " & CStr(pRecordset.RecordCount)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.Adodc2_MoveComplete(adReason,pError,adStatus,pRecordset)", Array(adReason, _
         pError, adStatus, pRecordset), EA_NORERAISE
    HandleError
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo errHandler
    If flgLoaded Then
        LoadNewData pRecordset
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.Adodc1_MoveComplete(adReason,pError,adStatus,pRecordset)", Array(adReason, _
         pError, adStatus, pRecordset), EA_NORERAISE
    HandleError
End Sub

Sub LoadNewData(pRecordset As ADODB.Recordset)
    On Error GoTo errHandler
    flgLoading = True
    If pRecordset.EOF = True Then Exit Sub
    
    Me.Adodc2.Recordset.Filter = "TP_Acno = '" & pRecordset.Fields("TP_Acno") & "'"
    If Adodc2.Recordset.EOF Then Exit Sub
    Me.Adodc2.Recordset.Requery
      flgLoading = False

    txtNameN = pRecordset.Fields("TP_Name")
    txtInitialsN = FNS(pRecordset.Fields("TP_Initials"))
    txtTitleN = FNS(pRecordset.Fields("TP_Title"))
    txtAddresseeN = FNS(pRecordset.Fields("ADD_Addressee"))
    txtAL1N = FNS(pRecordset.Fields("ADD_L1"))
    txtAL2N = FNS(pRecordset.Fields("ADD_L2"))
    txtAL3N = FNS(pRecordset.Fields("ADD_L3"))
    txtAL4N = FNS(pRecordset.Fields("ADD_L4"))
    txtAL5N = FNS(pRecordset.Fields("ADD_L5"))
    txtAL6N = FNS(pRecordset.Fields("ADD_L6"))
    txtPCodeN = FNS(pRecordset.Fields("ADD_PCode"))
    txtCountryN = tlCTRY.Item(FNN(Adodc2.Recordset.Fields("ADD_CNTRY_ID")))
    txtPhoneN = FNS(pRecordset.Fields("ADD_Phone"))
    txtCellN = FNS(pRecordset.Fields("TP_Cell"))
    txtEmailN = FNS(pRecordset.Fields("ADD_EmAIL"))
    txtFaxN = FNS(pRecordset.Fields("ADD_FAX"))
    txtIDNUMN = FNS(pRecordset.Fields("TP_IDNUM"))
    txtPhone2N = FNS(pRecordset.Fields("ADD_BUSPHONE"))
    
    
    If Not Adodc2.Recordset.EOF Then
        Me.cboCNTRY.Text = FNS(tlCTRY.Item(FNN(Adodc2.Recordset.Fields("ADD_CNTRY_ID"))))
    Else
        txtNameO = ""
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
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.LoadNewData(pRecordset)", pRecordset
    Exit Sub
    Resume
End Sub


Private Sub cboCNTRY_Click()
    On Error GoTo errHandler
'    tlCTRY.Item (FNN(Adodc2.Recordset.Fields("ADD_CNTRY_ID")))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cboCNTRY_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddressee_Click()
    On Error GoTo errHandler
    txtAddresseeO = txtAddresseeN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAddressee_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL1_Click()
    On Error GoTo errHandler
    txtAL1O = txtAL1N
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL2_Click()
    On Error GoTo errHandler
    txtAL2O = txtAL2N

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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL4_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL5_Click()
    On Error GoTo errHandler
    txtAL5O = txtAL5N

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL5_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAL6_Click()
    On Error GoTo errHandler
    txtAL6O = txtAL6N

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdAL6_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdIDNum_Click()
    On Error GoTo errHandler
    txtIDNUMO = txtIDNUMN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdIDNum_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCell_Click()
    On Error GoTo errHandler
    txtCellO = txtCellN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdCell_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPhone2_Click()
    On Error GoTo errHandler
    txtPhone2O = txtPhone2N
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdPhone2_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdFax_Click()
    On Error GoTo errHandler
    txtFaxO = txtFaxN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdFax_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEmail_Click()
    On Error GoTo errHandler
    txtEMailO = txtEmailN
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

'Private Sub cmdDefaultLayout_Click()
'SetDefaultWidths
'End Sub
'
'Private Sub cmdSaveLayout_Click()
'    SaveLayout G, Me.Name & "G"
'End Sub





Private Sub cmdInitials_Click()
    On Error GoTo errHandler
    txtInitialsO = txtInitialsN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdInitials_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdName_Click()
    On Error GoTo errHandler
    txtNameO = txtNameN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdName_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdPCode_Click()
    On Error GoTo errHandler
    txtPCodeO = txtPCodeN

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdPCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPhone_Click()
    On Error GoTo errHandler
    txtPhoneO = txtPhoneN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdPhone_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo errHandler
Dim bkmark
    Me.Adodc1.Recordset.Delete
    Me.Adodc1.Recordset.Update
    G.Refresh
    Me.Adodc1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTitle_Click()
    On Error GoTo errHandler
    txtTitleO = txtTitleN
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

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    LoadCombo cboCNTRY, tlCTRY
    Me.Width = 11600
    Me.Height = 9300
    Me.top = 160
    Me.left = 200
    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select * FROM tTP_Forapproval JOIN tADD_ForApproval ON ADD_TP_ID = TP_ID ORDER BY TP_NAME"
    Me.Adodc1.ConnectionString = oPC.CO.ConnectionString
    Me.Adodc1.CursorType = adOpenDynamic
    G.DataSource = Me.Adodc1
    G.ExtendRightColumn = False
    Me.Adodc2.CommandType = adCmdText
    Me.Adodc2.RecordSource = "Select * FROM tTP JOIN tADD ON ADD_TP_ID = TP_ID"  'Where TP_Acno = '" & Adodc1.Recordset.Fields("TP_ACNO") & "'"
    Me.Adodc2.ConnectionString = oPC.CO.ConnectionString
    Me.Adodc2.Refresh
    flgLoaded = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.Form_Load", , EA_NORERAISE
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
    txtAddresseeO.BackColor = IIf(txtAddresseeO = FNS(Adodc2.Recordset.Fields("ADD_Addressee").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAddresseeO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL1O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAL1O.BackColor = IIf(txtAL1O = FNS(Adodc2.Recordset.Fields("ADD_L1").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL1O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL2O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAL2O.BackColor = IIf(txtAL2O = FNS(Adodc2.Recordset.Fields("ADD_L2").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL2O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL3O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAL3O.BackColor = IIf(txtAL3O = FNS(Adodc2.Recordset.Fields("ADD_L3").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL3O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL4O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAL4O.BackColor = IIf(txtAL4O = FNS(Adodc2.Recordset.Fields("ADD_L4").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL4O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL5O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAL5O.BackColor = IIf(txtAL5O = FNS(Adodc2.Recordset.Fields("ADD_L5").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL5O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAL6O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAL6O.BackColor = IIf(txtAL6O = FNS(Adodc2.Recordset.Fields("ADD_L6").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtAL6O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCellO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtCellO.BackColor = IIf(txtCellO = FNS(Adodc2.Recordset.Fields("TP_CELL").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtCellO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInitialsO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtInitialsO.BackColor = IIf(txtInitialsO = FNS(Adodc2.Recordset.Fields("TP_INITIALS").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtInitialsO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtNameO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNameO.BackColor = IIf(txtNameO = FNS(Adodc2.Recordset.Fields("TP_Name").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtNameO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPCodeO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPCodeO.BackColor = IIf(txtPCodeO = FNS(Adodc2.Recordset.Fields("ADD_PCODE").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtPCodeO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhoneO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPhoneO.BackColor = IIf(txtPhoneO = FNS(Adodc2.Recordset.Fields("ADD_PHONE").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtPhoneO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitleO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtTitleO.BackColor = IIf(txtTitleO = FNS(Adodc2.Recordset.Fields("TP_TITLE").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtTitleO_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtPhone2O_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPhone2O.BackColor = IIf(txtPhone2O = FNS(Adodc2.Recordset.Fields("ADD_BUSPHONE").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtPhone2O_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtIDNUMO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtIDNUMO.BackColor = IIf(txtIDNUMO = FNS(Adodc2.Recordset.Fields("TP_IDNUM").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtIDNUMO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFaxO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtFaxO.BackColor = IIf(txtFaxO = FNS(Adodc2.Recordset.Fields("ADD_FAX").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtFaxO_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEMailO_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEMailO.BackColor = IIf(txtEMailO = FNS(Adodc2.Recordset.Fields("ADD_EMAIL").OriginalValue), vbWhite, &HC0C0FF)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPApprove.txtEMailO_Change", , EA_NORERAISE
    HandleError
End Sub

