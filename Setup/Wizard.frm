VERSION 5.00
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Papyrus II configuration wizard"
   ClientHeight    =   5220
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   7755
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0089524E&
   Icon            =   "Wizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7755
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Papyrus II configuration introduction"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   0
      Left            =   -10000
      TabIndex        =   6
      Tag             =   "1000"
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox chkShowIntro 
         Caption         =   "chkShowIntro"
         Height          =   315
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   18
         Tag             =   "1002"
         Top             =   4065
         Width           =   3810
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0089524E&
         Height          =   675
         Index           =   0
         Left            =   465
         TabIndex        =   7
         Tag             =   "1001"
         Top             =   1140
         Width           =   6405
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1230
         Index           =   0
         Left            =   105
         Picture         =   "Wizard.frx":0442
         Stretch         =   -1  'True
         Top             =   3165
         Width           =   3180
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   1
      Left            =   -10000
      TabIndex        =   8
      Tag             =   "2000"
      Top             =   0
      Width           =   7155
      Begin VB.TextBox txtBankingDetails 
         Height          =   735
         Left            =   2970
         MultiLine       =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "Your banking details, used on invoices."
         Top             =   3090
         Width           =   2295
      End
      Begin VB.TextBox txtCode 
         Height          =   360
         Left            =   2970
         MaxLength       =   2
         TabIndex        =   27
         ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
         Top             =   1140
         Width           =   1335
      End
      Begin VB.TextBox txtVATNumber 
         Height          =   360
         Left            =   135
         MaxLength       =   40
         TabIndex        =   26
         ToolTipText     =   "Your V.A.T. number, if applicable."
         Top             =   3795
         Width           =   2400
      End
      Begin VB.TextBox txtRegistrationNumber 
         Height          =   360
         Left            =   135
         MaxLength       =   40
         TabIndex        =   25
         ToolTipText     =   "Your CC number or company registration number, if applicable."
         Top             =   3090
         Width           =   2400
      End
      Begin VB.TextBox txtStreetAddress 
         Height          =   870
         Left            =   2955
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1890
         Width           =   2295
      End
      Begin VB.TextBox txtPostalAddress 
         Height          =   870
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1890
         Width           =   2385
      End
      Begin VB.CheckBox chkUsetest 
         Caption         =   "Apply these settings to the test database"
         Height          =   360
         Left            =   165
         TabIndex        =   22
         ToolTipText     =   "All settings that follow will apply only to the installed test database"
         Top             =   510
         Width           =   3630
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   360
         Left            =   135
         TabIndex        =   20
         ToolTipText     =   "Your business name, used on invoices etc."
         Top             =   1140
         Width           =   2385
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Banking details"
         Height          =   285
         Left            =   3015
         TabIndex        =   34
         Tag             =   "2014"
         Top             =   2850
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Business V.A.T. number"
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Tag             =   "2013"
         Top             =   3555
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Business registration number"
         Height          =   285
         Left            =   135
         TabIndex        =   32
         Tag             =   "2012"
         Top             =   2850
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Street address"
         Height          =   285
         Left            =   2985
         TabIndex        =   31
         Tag             =   "2011"
         Top             =   1650
         Width           =   1800
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Postal address"
         Height          =   285
         Left            =   135
         TabIndex        =   30
         Tag             =   "2010"
         Top             =   1650
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Business short code"
         Height          =   285
         Left            =   3045
         TabIndex        =   29
         Tag             =   "2009"
         Top             =   885
         Width           =   1800
      End
      Begin VB.Label lblCompanyname 
         BackStyle       =   0  'Transparent
         Caption         =   "Business name"
         Height          =   285
         Left            =   135
         TabIndex        =   21
         Tag             =   "2008"
         Top             =   885
         Width           =   1800
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0089524E&
         Height          =   390
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Tag             =   "2001"
         Top             =   75
         Width           =   6855
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Tag             =   "2002"
      Top             =   0
      Width           =   7155
      Begin VB.Frame frStore1 
         Caption         =   "Store 3"
         Height          =   3300
         Index           =   2
         Left            =   4860
         TabIndex        =   52
         Top             =   1140
         Width           =   2205
         Begin VB.TextBox txtStore1Name 
            Height          =   285
            Index           =   2
            Left            =   135
            MaxLength       =   40
            TabIndex        =   56
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   480
            Width           =   1395
         End
         Begin VB.TextBox txtStore1Code 
            Height          =   285
            Index           =   2
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   55
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   480
            Width           =   390
         End
         Begin VB.TextBox txtStore1StreetAddress 
            Height          =   870
            Index           =   2
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   54
            Top             =   2190
            Width           =   1830
         End
         Begin VB.TextBox txtStore1PostalAddress 
            Height          =   870
            Index           =   2
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   1065
            Width           =   1830
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Name and short code"
            Height          =   285
            Index           =   2
            Left            =   195
            TabIndex        =   59
            Tag             =   "2016"
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Street address"
            Height          =   285
            Index           =   2
            Left            =   195
            TabIndex        =   58
            Tag             =   "2011"
            Top             =   1980
            Width           =   1740
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal address"
            Height          =   285
            Index           =   2
            Left            =   195
            TabIndex        =   57
            Tag             =   "2010"
            Top             =   825
            Width           =   1740
         End
      End
      Begin VB.Frame frStore1 
         Caption         =   "Store 2"
         Height          =   3300
         Index           =   1
         Left            =   2535
         TabIndex        =   44
         Top             =   1140
         Width           =   2205
         Begin VB.TextBox txtStore1PostalAddress 
            Height          =   870
            Index           =   1
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   1065
            Width           =   1830
         End
         Begin VB.TextBox txtStore1StreetAddress 
            Height          =   870
            Index           =   1
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   2190
            Width           =   1830
         End
         Begin VB.TextBox txtStore1Code 
            Height          =   285
            Index           =   1
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   46
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   480
            Width           =   390
         End
         Begin VB.TextBox txtStore1Name 
            Height          =   285
            Index           =   1
            Left            =   135
            MaxLength       =   40
            TabIndex        =   45
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal address"
            Height          =   285
            Index           =   1
            Left            =   195
            TabIndex        =   51
            Tag             =   "2010"
            Top             =   825
            Width           =   1740
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Street address"
            Height          =   285
            Index           =   1
            Left            =   195
            TabIndex        =   50
            Tag             =   "2011"
            Top             =   1980
            Width           =   1740
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Name and short code"
            Height          =   285
            Index           =   1
            Left            =   195
            TabIndex        =   49
            Tag             =   "2016"
            Top             =   240
            Width           =   1740
         End
      End
      Begin VB.Frame frStore1 
         Caption         =   "Store 1"
         Height          =   3300
         Index           =   0
         Left            =   225
         TabIndex        =   36
         Top             =   1140
         Width           =   2205
         Begin VB.TextBox txtStore1Name 
            Height          =   285
            Index           =   0
            Left            =   135
            MaxLength       =   40
            TabIndex        =   42
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   480
            Width           =   1395
         End
         Begin VB.TextBox txtStore1Code 
            Height          =   285
            Index           =   0
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   41
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   480
            Width           =   390
         End
         Begin VB.TextBox txtStore1StreetAddress 
            Height          =   870
            Index           =   0
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   2190
            Width           =   1830
         End
         Begin VB.TextBox txtStore1PostalAddress 
            Height          =   870
            Index           =   0
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   1065
            Width           =   1830
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Name and short code"
            Height          =   285
            Index           =   0
            Left            =   195
            TabIndex        =   43
            Tag             =   "2016"
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Street address"
            Height          =   285
            Index           =   0
            Left            =   195
            TabIndex        =   40
            Tag             =   "2011"
            Top             =   1980
            Width           =   1740
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal address"
            Height          =   285
            Index           =   0
            Left            =   195
            TabIndex        =   39
            Tag             =   "2010"
            Top             =   825
            Width           =   1740
         End
      End
      Begin VB.CheckBox chkOneStore 
         Caption         =   "We have only one store and its details are the same as those for the business itself."
         Height          =   195
         Left            =   510
         TabIndex        =   35
         Top             =   870
         Width           =   6405
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0089524E&
         Height          =   645
         Index           =   2
         Left            =   405
         TabIndex        =   11
         Tag             =   "2003"
         Top             =   120
         Width           =   6450
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Index           =   3
      Left            =   -10000
      TabIndex        =   12
      Tag             =   "2004"
      Top             =   15
      Width           =   7155
      Begin VB.Frame Frame4 
         Height          =   795
         Left            =   90
         TabIndex        =   96
         Top             =   3765
         Width           =   7050
         Begin VB.TextBox txtFullname 
            Height          =   285
            Index           =   4
            Left            =   120
            MaxLength       =   40
            TabIndex        =   101
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1425
         End
         Begin VB.TextBox txtShortname 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1635
            MaxLength       =   3
            TabIndex        =   100
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txtSignature 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   2595
            MaxLength       =   40
            TabIndex        =   99
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1230
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Supervisor"
            Height          =   270
            Left            =   4080
            TabIndex        =   98
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Operator"
            Height          =   270
            Left            =   5460
            TabIndex        =   97
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Full name"
            Height          =   285
            Index           =   6
            Left            =   180
            TabIndex        =   104
            Tag             =   "30002"
            Top             =   165
            Width           =   1305
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Short code"
            Height          =   285
            Index           =   16
            Left            =   1665
            TabIndex        =   103
            Tag             =   "30003"
            Top             =   165
            Width           =   810
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Email signature"
            Height          =   285
            Index           =   17
            Left            =   2685
            TabIndex        =   102
            Tag             =   "30004"
            Top             =   165
            Width           =   1200
         End
      End
      Begin VB.Frame Frame3 
         Height          =   795
         Left            =   90
         TabIndex        =   87
         Top             =   2985
         Width           =   7050
         Begin VB.OptionButton Option6 
            Caption         =   "Operator"
            Height          =   270
            Left            =   5460
            TabIndex        =   92
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Supervisor"
            Height          =   270
            Left            =   4080
            TabIndex        =   91
            Top             =   375
            Width           =   1140
         End
         Begin VB.TextBox txtSignature 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   2595
            MaxLength       =   40
            TabIndex        =   90
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1230
         End
         Begin VB.TextBox txtShortname 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1635
            MaxLength       =   3
            TabIndex        =   89
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txtFullname 
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   40
            TabIndex        =   88
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1425
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Email signature"
            Height          =   285
            Index           =   13
            Left            =   2685
            TabIndex        =   95
            Tag             =   "30004"
            Top             =   165
            Width           =   1200
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Short code"
            Height          =   285
            Index           =   14
            Left            =   1665
            TabIndex        =   94
            Tag             =   "30003"
            Top             =   165
            Width           =   810
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Full name"
            Height          =   285
            Index           =   15
            Left            =   180
            TabIndex        =   93
            Tag             =   "30002"
            Top             =   165
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   90
         TabIndex        =   78
         Top             =   2205
         Width           =   7050
         Begin VB.TextBox txtFullname 
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   40
            TabIndex        =   83
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1425
         End
         Begin VB.TextBox txtShortname 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1635
            MaxLength       =   3
            TabIndex        =   82
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txtSignature 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   2595
            MaxLength       =   40
            TabIndex        =   81
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1230
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Supervisor"
            Height          =   270
            Left            =   4080
            TabIndex        =   80
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Operator"
            Height          =   270
            Left            =   5460
            TabIndex        =   79
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Full name"
            Height          =   285
            Index           =   10
            Left            =   180
            TabIndex        =   86
            Tag             =   "30002"
            Top             =   165
            Width           =   1305
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Short code"
            Height          =   285
            Index           =   11
            Left            =   1665
            TabIndex        =   85
            Tag             =   "30003"
            Top             =   165
            Width           =   810
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Email signature"
            Height          =   285
            Index           =   12
            Left            =   2685
            TabIndex        =   84
            Tag             =   "30004"
            Top             =   150
            Width           =   1200
         End
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   75
         TabIndex        =   69
         Top             =   1425
         Width           =   7050
         Begin VB.TextBox txtFullname 
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   40
            TabIndex        =   74
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1425
         End
         Begin VB.TextBox txtShortname 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1635
            MaxLength       =   3
            TabIndex        =   73
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txtSignature 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2595
            MaxLength       =   40
            TabIndex        =   72
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1230
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Supervisor"
            Height          =   270
            Left            =   4080
            TabIndex        =   71
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Operator"
            Height          =   270
            Left            =   5460
            TabIndex        =   70
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Full name"
            Height          =   285
            Index           =   9
            Left            =   180
            TabIndex        =   77
            Tag             =   "30002"
            Top             =   165
            Width           =   1305
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Short code"
            Height          =   285
            Index           =   8
            Left            =   1665
            TabIndex        =   76
            Tag             =   "30003"
            Top             =   165
            Width           =   810
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Email signature"
            Height          =   285
            Index           =   7
            Left            =   2685
            TabIndex        =   75
            Tag             =   "30004"
            Top             =   165
            Width           =   1200
         End
      End
      Begin VB.Frame frSM 
         Height          =   795
         Left            =   75
         TabIndex        =   60
         Top             =   645
         Width           =   7050
         Begin VB.OptionButton Option1 
            Caption         =   "Operator"
            ForeColor       =   &H80000017&
            Height          =   270
            Left            =   5460
            TabIndex        =   68
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton optSupervisor 
            Caption         =   "Supervisor"
            ForeColor       =   &H80000017&
            Height          =   270
            Left            =   4080
            TabIndex        =   67
            Top             =   375
            Width           =   1140
         End
         Begin VB.TextBox txtSignature 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2595
            MaxLength       =   40
            TabIndex        =   65
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1230
         End
         Begin VB.TextBox txtShortname 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1635
            MaxLength       =   3
            TabIndex        =   63
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txtFullname 
            Height          =   285
            Index           =   3
            Left            =   120
            MaxLength       =   40
            TabIndex        =   61
            ToolTipText     =   "A one or two letter prefix that will be used to identify this business"
            Top             =   375
            Width           =   1425
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Email signature"
            ForeColor       =   &H80000015&
            Height          =   285
            Index           =   5
            Left            =   2685
            TabIndex        =   66
            Tag             =   "30004"
            Top             =   165
            Width           =   1200
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Short code"
            ForeColor       =   &H80000017&
            Height          =   285
            Index           =   4
            Left            =   1665
            TabIndex        =   64
            Tag             =   "30003"
            Top             =   165
            Width           =   810
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Full name"
            ForeColor       =   &H80000017&
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   62
            Tag             =   "30002"
            Top             =   165
            Width           =   1305
         End
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "sd sdf sd "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0089524E&
         Height          =   570
         Index           =   3
         Left            =   135
         TabIndex        =   13
         Tag             =   "2005"
         Top             =   30
         Width           =   6975
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   4
      Left            =   -10000
      TabIndex        =   14
      Tag             =   "2006"
      Top             =   0
      Width           =   7155
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   4
         Left            =   2700
         TabIndex        =   15
         Tag             =   "2007"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1995
         Index           =   4
         Left            =   210
         Picture         =   "Wizard.frx":31806
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Finished!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   5
      Left            =   -10000
      TabIndex        =   16
      Tag             =   "3000"
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox chkSaveSettings 
         Caption         =   "chkSaveSettings"
         Height          =   552
         Left            =   3210
         MaskColor       =   &H00000000&
         TabIndex        =   19
         Tag             =   "3003"
         Top             =   1650
         Width           =   3552
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   5
         Left            =   3210
         TabIndex        =   17
         Tag             =   "3001"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   5
         Left            =   210
         Picture         =   "Wizard.frx":360DC
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   4650
      Width           =   7755
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   5925
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   4545
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   2250
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   108
         X2              =   7012
         Y1              =   24
         Y2              =   24
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 6

Const RES_ERROR_MSG = 30000

'BASE VALUE FOR HELP FILE FOR THIS WIZARD:
Const HELP_BASE = 1000
Const HELP_FILE = "MYWIZARD.HLP"

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_FINISH = 5

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Papyrus II configuration wizard"
Const INTRO_KEY = "IntroductionScreen"
Const SHOW_INTRO = "ShowIntro"
Const TOPIC_TEXT = "<TOPIC_TEXT>"
'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

Private Sub chkShowIntro_Click()
    If chkShowIntro.Value Then
        SaveSetting APP_CATEGORY, WIZARD_NAME, INTRO_KEY, SHOW_INTRO
    Else
        SaveSetting APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString
    End If
End Sub


Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    
    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)
            rc = WinHelp(Me.hwnd, HELP_FILE, HELP_CONTEXT, lHelpTopic)
        
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep + 1
            SetStep nAltStep, DIR_NEXT
          
        Case BTN_FINISH
            'wizard creation code goes here
            ConnectToDatabase
            PostChanges
            
            Unload Me
            
            If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
                frmConfirm.Show vbModal
            End If
        
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_HELP
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False
    
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    'Load All string info for Form
    LoadResStrings Me
    
    'Determine 1st Step:
    If GetSetting(APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString) = SHOW_INTRO Then
        chkShowIntro.Value = vbChecked
        SetStep 1, DIR_NEXT
    Else
        SetStep 0, DIR_NONE
    End If
End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Select Case nStep
        Case STEP_INTRO
      
        Case STEP_1
      
        Case STEP_2
            
        
        Case STEP_3
      
        Case STEP_4
            mbFinishOK = False
      
        Case STEP_FINISH
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
    If chkSaveSettings.Value = vbChecked Then
      
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
      
    End If
  
    If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub



Private Sub chkUsetest_Click()
    bUseTest = (chkUsetest = 1)
End Sub

Sub PostChanges()
Dim strSQL As String

'Post company details
    strSQL = "UPDATE tCompany SET COMP_NAME = '" & Trim(Me.txtCompanyName) & "'," _
    & "COMP_STREETADD = '" & Trim(Me.txtStreetAddress) & "'," _
    & "COMP_POSTALADD = '" & Trim(Me.txtPostalAddress) & "'," _
    & "COMP_REGISTRATIONNUMBER = '" & Trim(Me.txtRegistrationNumber) & "'," _
    & "COMP_VATNUMBER = '" & Trim(Me.txtVATNumber) & "'," _
    & "COMP_CODE = '" & Left(Trim(Me.txtCode), 2) & "'," _
    & "COMP_BANKDETAILS = '" & Trim(Me.txtBankingDetails) & "' FROM tCOMPANY JOIN tConfiguration ON COMP_ID = CF_DEFAULTCOMPANYID"

    oCnn.Execute strSQL
    
    'Post store details
    For i = 0 To 2
        If txtStore1Name(i) > "" Then
            PostStore txtStore1Name(i), txtStore1Code(i), txtStore1PostalAddress(i), txtStore1PostalAddress(i)
        End If
    Next i
    'Post staffmember details
    
    'Post email details
    
    'Post configuration details
End Sub

Public Sub AutoSelect(ctl As Control)
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
End Sub



Private Sub txtBankingDetails_GotFocus()
    AutoSelect Controls("txtBankingDetails")
End Sub

Private Sub txtCode_GotFocus()
    AutoSelect Controls("txtCode")
End Sub

Private Sub txtCompanyName_GotFocus()
    AutoSelect Controls("txtCompanyName")
End Sub

Private Sub txtPostalAddress_GotFocus()
    AutoSelect Controls("txtPostalAddress")
End Sub

Private Sub txtRegistrationNumber_GotFocus()
    AutoSelect Controls("txtRegistrationNumber")
End Sub



Private Sub txtStreetAddress_GotFocus()
    AutoSelect Controls("txtStreetAddress")
End Sub

Private Sub txtVATNumber_GotFocus()
    AutoSelect Controls("txtVATNumber")
End Sub

Private Sub chkOneStore_Click()
    If chkOneStore = 1 Then
        ClearStore 1
        ClearStore 2
        EnableAllStoreFrames False
    Else
        EnableAllStoreFrames True
        If txtStore1Name(1) = "" Then
            CopyStoreDetailsFromCompany
        End If
    End If
End Sub
Private Sub EnableAllStoreFrames(Enable As Boolean)
Dim i As Integer

    For i = 0 To 2
        txtStore1Name(i).Enabled = Enable
        txtStore1PostalAddress(i).Enabled = Enable
        txtStore1StreetAddress(i).Enabled = Enable
        txtStore1Code(i).Enabled = Enable
    Next
End Sub

Private Sub ClearStore(StoreNumber As Integer)
    txtStore1Name(StoreNumber) = ""
    txtStore1PostalAddress(StoreNumber) = ""
    txtStore1StreetAddress(StoreNumber) = ""
    txtStore1Code(StoreNumber) = ""
End Sub

Private Sub CopyStoreDetailsFromCompany()
    txtStore1Name(0) = txtCompanyName
    txtStore1PostalAddress(0) = txtPostalAddress
    txtStore1StreetAddress(0) = txtStreetAddress
    txtStore1Code(0) = txtCode
End Sub

Sub PostStore(pName As String, pShortCode As String, pPostal As String, pStreet As String)
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
    Set cmd = New ADODB.Command
    cmd.CommandText = "PostStore_FromConfiguration"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, pName)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ShortCode", adVarChar, adParamInput, 5, pShortCode)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PostalAddress", adVarChar, adParamInput, 250, pPostal)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@StreetAddress", adVarChar, adParamInput, 250, pStreet)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oCnn
    cmd.Execute
    
    Set cmd = Nothing

End Sub


