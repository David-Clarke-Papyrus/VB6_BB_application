VERSION 5.00
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10920
   Begin VB.CheckBox chkExSales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude from sales reporting"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   4035
      TabIndex        =   44
      Top             =   6975
      Width           =   2550
   End
   Begin VB.ComboBox cboStores 
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
      Left            =   4035
      TabIndex        =   8
      Text            =   "cboStores"
      Top             =   6540
      Width           =   2100
   End
   Begin VB.TextBox txtIDNum 
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
      Left            =   8430
      TabIndex        =   5
      Top             =   1425
      Width           =   2205
   End
   Begin VB.CheckBox chkVATable 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Pays V.A.T"
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
      Left            =   150
      TabIndex        =   40
      Top             =   6660
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.CheckBox chkTemporary 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Temporary customer"
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
      Left            =   90
      TabIndex        =   39
      Top             =   6885
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   6465
      TabIndex        =   4
      Top             =   1455
      Width           =   1065
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
      Left            =   3465
      TabIndex        =   1
      Top             =   315
      Width           =   1965
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
      Left            =   5805
      TabIndex        =   2
      Top             =   300
      Width           =   1020
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
      Left            =   7395
      TabIndex        =   3
      Top             =   300
      Width           =   3180
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
      Left            =   7995
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Customer group membership"
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
      Left            =   6285
      TabIndex        =   34
      Top             =   1860
      Width           =   4380
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1050
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
         TabIndex        =   10
         Top             =   345
         Width           =   1305
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
         TabIndex        =   9
         Top             =   375
         Width           =   2745
      End
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   795
         Width           =   2700
      End
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
      Height          =   4470
      Left            =   135
      TabIndex        =   32
      Top             =   885
      Width           =   6060
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3165
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3855
         Width           =   930
      End
      Begin VB.CommandButton cmdRemove 
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
         Height          =   405
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3855
         Width           =   945
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
         Height          =   405
         Left            =   5070
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3855
         Width           =   870
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         DragIcon        =   "frmCustomer.frx":0000
         Height          =   3525
         Left            =   120
         OleObjectBlob   =   "frmCustomer.frx":0442
         TabIndex        =   17
         Top             =   300
         Width           =   5805
      End
      Begin CoolButtonControl.CoolButton cbBillTo 
         Height          =   300
         Left            =   2445
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3855
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bill"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cmdDefaultAddress 
         Height          =   300
         Left            =   1680
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3855
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Appro"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbDelTo 
         Height          =   300
         Left            =   915
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3840
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Deliver"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbOrderTo 
         Height          =   300
         Left            =   135
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3855
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Order"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.Label lblRecords 
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   41
         Top             =   4170
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Interest group membership"
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
      Left            =   6270
      TabIndex        =   31
      Top             =   4185
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
         TabIndex        =   16
         TabStop         =   0   'False
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   795
         Width           =   2700
      End
   End
   Begin VB.TextBox txtDefaultDiscount 
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
      Left            =   150
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7140
      Visible         =   0   'False
      Width           =   720
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
      Height          =   885
      Left            =   750
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5490
      Width           =   5385
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
      Left            =   8700
      Picture         =   "frmCustomer.frx":3899
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6375
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
      Left            =   9675
      Picture         =   "frmCustomer.frx":3E23
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
      Width           =   990
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
      Top             =   330
      Width           =   3000
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Originating store"
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
      Left            =   2430
      TabIndex        =   43
      Top             =   6570
      Width           =   1560
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "I.D. Num."
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
      Left            =   8745
      TabIndex        =   42
      Top             =   1185
      Width           =   1065
   End
   Begin VB.Label Label3 
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
      Left            =   6345
      TabIndex        =   38
      Top             =   1185
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   300
      Left            =   7350
      TabIndex        =   37
      Top             =   45
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "First name (if person)"
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
      Left            =   3510
      TabIndex        =   36
      Top             =   45
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Title (if person)"
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
      Left            =   5700
      TabIndex        =   35
      Top             =   45
      Width           =   1290
   End
   Begin VB.Label Label4 
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
      Left            =   165
      TabIndex        =   33
      Top             =   5445
      Width           =   465
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Default discount"
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
      Left            =   930
      TabIndex        =   30
      Top             =   7185
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2460
      X2              =   300
      Y1              =   150
      Y2              =   810
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   765
      Left            =   6225
      TabIndex        =   28
      Top             =   6270
      Width           =   2445
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
      TabIndex        =   27
      Top             =   75
      Width           =   735
   End
   Begin VB.Menu mnuACtions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete customer"
      End
   End
End
Attribute VB_Name = "frmCustomer"
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


Public Property Get EMail() As String
    EMail = strEMail
End Property


Private Sub cboStores_Click()
If flgLoading Then Exit Sub
    oCust.SetStoreID oPC.Configuration.Stores_tl.Key(cboStores)
End Sub

Private Sub cboStores_Validate(Cancel As Boolean)
If flgLoading Then Exit Sub
    oCust.SetStoreID oPC.Configuration.Stores_tl.Key(cboStores)
End Sub

Private Sub chkExSales_Click()
oCust.ExcludeFromSales = IIf(Me.chkExSales = 1, True, False)
End Sub

'Private Sub cboCT_Click()
'    If flgLoading Then Exit Sub
'    oCust.CustomerTypeID = oCust.CustomerTypes_tl.Key(cboCT)
'End Sub

'Private Sub chkTemp_Click()
'    If flgLoading Then Exit Sub
'    oCust.CanBeDeleted = (Me.chkTemp = 1)
'
'End Sub

Private Sub cmdAddCC_Click()
    On Error GoTo errHandler
Dim oCC As New a_IG
    If flgLoading Then Exit Sub
    If cboCC = "" Then Exit Sub
    Set oCC = oCust.CustomerTypes.Add
 '   If CustID = 0 Then Exit Sub
    oCC.BeginEdit
   ' oCC.TPID = oCust.ID
    oCC.IGID = oCust.CustomerTypes_tl.Key(cboCC)
    oCC.Description = cboCC
    oCC.ApplyEdit
    
'    oCust.CustomerTypes.ApplyEdit
'    oCust.CustomerTypes.BeginEdit
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

Private Sub cmdRemoveCC_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbCC = "" Then Exit Sub
    oCust.CustomerTypes.Remove oCust.CustomerTypes.Key(Me.lbCC)
    cboCC.AddItem Me.lbCC
    cboCC.ListIndex = 0
    LoadTPCCs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyalty.cmdRemoveCC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkVatable_Click()
    If flgLoading Then Exit Sub
    oCust.VATable = (chkVATable = 1)
End Sub

Private Sub cmdAdd_Click()
Dim frm As frmAddress
Dim oAdd As a_Address
    If flgLoading Then Exit Sub
    Set frm = New frmAddress
    Set oAdd = oCust.Addresses.Add
    oAdd.BeginEdit
    oAdd.SetAddressee oCust.Title & " " & oCust.Initials & " " & oCust.Name
    frm.Component oAdd
    frm.Show vbModal
    LoadArray
    LoadIGs
End Sub


Private Sub cmdAddIG_Click()
On Error GoTo Errh
Dim oIG As a_IG
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
Errh:
    If Err = -2147220949 Then
        MsgBox "This item is already selected for this customer"
        Exit Sub
    End If
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
    ErrorIn "frmCustomer.cmdRemoveIG_Click", , EA_NORERAISE
    HandleError
End Sub



'Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'    oCust.Addresses(XA(G1.Bookmark, 9)) = G1.Text
'End Sub

Private Sub mnuDel_Click()
'Dim ocInv As New c_Invoices
'Dim bRecsreturned As Boolean
'    If flgLoading Then Exit Sub
'    ocInv.Load bRecsreturned, oCust.ID
'    If ocInv.Count > 0 Then
'        MsgBox "There are invoices stored for this customer. You cannot delete it.", vbInformation, "Action denied"
'        Exit Sub
'    End If
'    Set ocInv = Nothing
'    MsgBox "Note to david : Check customer orders also"
'    Me.LinCancel.Visible = True
'    oCust.DeleteCustomer
End Sub

'Private Sub oCust_ApproAddressChanged()
'On Error Resume Next
'    Me.txtPhone = oCust.ApproAddress.Phone
'  '  LoadAddresses
'End Sub
'Private Sub LoadClassErrorsCollection()
''In order to report user-understandable messages, this class holds a collection of short message
''codes paired with full descriptive messages.
''The collection is loaded here
'    Set colClassErrors = New Collection
'    colClassErrors.Add "Every customer must ahve a name.", "Name"
'    colClassErrors.Add "Every customer must have at least one phone number", "Phone"
'End Sub
'Private Function TranslateErrors(ByVal pRawErrors As String) As String
''Takes the short (raw) error messages used within this class and translates them to a
''formatted string  (including vbCRLFs) with full error descriptions. The result
''can be used in a message box at the GUI level
'Dim strRule As String
'Dim strAllRules As String
'Dim iMarker As Integer
'Dim iStart As Integer
'    iMarker = 1
'    strAllRules = ""
'    If Len(pRawErrors) > 0 Then
'        iMarker = InStr(iMarker + 1, pRawErrors, ",")
'        If iMarker > 0 Then
'            strAllRules = colClassErrors(Left$(pRawErrors, iMarker - 1))
'        Else
'            strAllRules = colClassErrors(pRawErrors)
'        End If
'        Do Until iMarker = 0
'            iStart = iMarker + 1
'            iMarker = InStr(iStart, pRawErrors, ",")
'            If iMarker > 0 Then
'                strRule = colClassErrors(Mid$(pRawErrors, iStart, iMarker - iStart))
'            Else
'                strRule = colClassErrors(Mid$(pRawErrors, iStart))
'            End If
'
'            strAllRules = strAllRules & vbCrLf & strRule
'        Loop
'    End If
'    TranslateErrors = strAllRules
'End Function
Public Sub Component(pCust As a_Customer)
    Set oCust = pCust
'    oCust.BeginEdit
    Me.Caption = "Customer: " & oCust.Name
End Sub
Private Sub EnableOK(pOK As Boolean)
    cmdOK.Enabled = pOK
End Sub


Private Sub cmdCancel_Click()
    oCust.CancelEdit
    Unload Me
End Sub



Private Sub cmdOK_Click()
Dim lngResult As Long
    If flgLoading Then Exit Sub
    oCust.LookforDuplicates
    If oCust.CustomerIndexClashes = True Then
        MsgBox "This account number has already been used for another customer. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    oCust.ApplyEdit lngResult
    If lngResult = 0 Then
        Unload Me
    ElseIf lngResult = 22 Then
        MsgBox "You are trying to save a customer with duplicate values." & vbCrLf & "These are likely to be in the Acc No. field or in the address description fields.", , "Can't save"
    End If
End Sub


Private Sub Form_Load()
    flgLoading = True
    Me.top = 0
    Me.left = 50
    Me.Height = 8200
    Me.Width = 11000
    txtName = oCust.Name
    txtFN = oCust.Initials
    txtAcno = oCust.AcNo
    txtTitle = oCust.Title
    txtNote = oCust.Note
    txtIDNum = oCust.IDNum
    txtMobile = oCust.Mobile
    chkVATable = IIf(oCust.VATable, 1, 0)
    txtDefaultDiscount = oCust.DefaultDiscountF
    Me.chkExSales = IIf(oCust.ExcludeFromSales, 1, 0)
    Me.txtDefaultDiscount = oCust.DefaultDiscountF
    Me.chkTemporary = IIf(oCust.CanBeDeleted, 1, 0)
    
    oPC.Configuration.LoadStores_tl ""
    LoadCombo cboStores, oPC.Configuration.Stores_tl
    cboStores = oCust.StoreName

    LoadArray
    LoadIGs
    LoadTPIGs
    LoadCCs
    LoadTPCCs
    RestrictInterestGroups
    RestrictCustomerTypes

    oCust.GetStatus
    flgLoading = False
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
Private Sub RestrictCustomerTypes()
Dim oTPIG As a_IG
Dim i As Integer

    For Each oTPIG In oCust.CustomerTypes
        For i = cboCC.ListCount To 1 Step -1
            cboCC.ListIndex = i - 1
            If oTPIG.Description = cboCC Then
                cboCC.RemoveItem cboCC.ListIndex
            End If
        Next
    Next
    If cboCC.ListCount > 0 Then
        cboCC.ListIndex = 0
    Else
        cboCC.ListIndex = -1
    End If
End Sub


Private Sub LoadIGs()
    LoadCombo Me.cboIG, oCust.InterestGroupsActive_tl
End Sub
Private Sub LoadCCs()
    LoadCombo Me.cboCC, oCust.CustomerTypes_tl
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
Private Sub LoadTPIGs()
Dim oTPIG As a_IG
    With Me.lbIG
        .Clear
        For Each oTPIG In oCust.InterestGroups
            .AddItem oCust.InterestGroupsAll_tl.Item(CStr(oTPIG.IGID))
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_C_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
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
        Set Forms(0).frmMainCustomerPreview = New frmCustomerPreview
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

Private Sub oCust_Valid(strMsg As String)
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
End Sub






Private Sub txtDefaultDiscount_LostFocus()
    txtDefaultDiscount = oCust.DefaultDiscountF
End Sub

Private Sub txtDefaultDiscount_Validate(Cancel As Boolean)
    Cancel = Not oCust.SetDefaultDiscount(txtDefaultDiscount)
End Sub
'Private Sub txtPhone_LostFocus()
'    txtPhone = oCust.Phone
'End Sub

'Private Sub txtPhone_Validate(Cancel As Boolean)
'    Cancel = Not oCust.SetPhone(txtPhone)
'End Sub
'Private Sub txtPhone_Change()
'Dim intPos As Integer
'    On Error Resume Next
'    oCust.SetPhone (txtPhone)
'    If Err Then
'      Beep
'      intPos = txtPhone.SelStart
'      txtPhone = oCust.Phone
'      txtPhone.SelStart = intPos - 1
'    End If
'End Sub

Private Sub txtName_LostFocus()
    txtName = oCust.Name
End Sub
Private Sub txtName_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.SetName (txtName)
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oCust.Name
      txtName.SelStart = intPos - 1
    End If
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    Cancel = Not oCust.SetName(txtName)
End Sub


Private Sub txtIDNum_LostFocus()
    txtIDNum = oCust.IDNum
End Sub
Private Sub txtIDNum_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.SetIDNum (txtIDNum)
    txtIDNum.BackColor = SetBackground(oCust.IDNum = txtIDNum.Text)
    
End Sub
Private Sub txtIDNum_Validate(Cancel As Boolean)
    Cancel = Not oCust.SetIDNum(txtIDNum)
End Sub

Public Function SetBackground(pGood As Boolean) As Long
    If pGood Then
        SetBackground = &H80000005
    Else
        SetBackground = &H80000004
    End If
End Function






Private Sub txtAcno_LostFocus()
    txtAcno = oCust.AcNo
End Sub
Private Sub txtAcno_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.SetAcNO (txtAcno)
    If Err Then
      Beep
      intPos = txtAcno.SelStart
      txtAcno = oCust.AcNo
      txtAcno.SelStart = intPos - 1
    End If
End Sub
Private Sub txtAcno_Validate(Cancel As Boolean)
    Cancel = Not oCust.SetAcNO(txtAcno)
End Sub
Private Sub txtFN_LostFocus()
    txtFN = oCust.Initials
End Sub
Private Sub txtFN_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.SetInitials (txtFN)
    If Err Then
      Beep
      intPos = txtFN.SelStart
      txtFN = oCust.Initials
      txtFN.SelStart = intPos - 1
    End If
End Sub
Private Sub txtFN_Validate(Cancel As Boolean)
    Cancel = Not oCust.SetInitials(txtFN)
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    Cancel = Not oCust.setnote(txtNote)
End Sub
Private Sub txtNote_LostFocus()
    txtNote = oCust.Note
End Sub
Private Sub txtNote_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCust.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub

Private Sub txtTitle_LostFocus()
    txtTitle = oCust.Title
End Sub
Private Sub txtTitle_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.SetTitle (txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oCust.Title
      txtTitle.SelStart = intPos - 1
    End If
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    Cancel = Not oCust.SetTitle(txtTitle)
End Sub
Private Sub txtMobile_LostFocus()
    txtMobile = oCust.Mobile
End Sub
Private Sub txtMobile_Change()
Dim intPos As Integer
    On Error Resume Next
    oCust.SetMobile (txtMobile)
    If Err Then
      Beep
      intPos = txtMobile.SelStart
      txtMobile = oCust.Title
      txtMobile.SelStart = intPos - 1
    End If
End Sub
Private Sub txtMobile_Validate(Cancel As Boolean)
    If txtMobile = "" Then Exit Sub
    txtMobile = PhoneFormat(txtMobile, oPC.DefaultAreaCode)
    Cancel = Not oCust.SetMobile(txtMobile)
End Sub


Private Sub LoadArray()
'Dim objItem As d_Customer
Dim lngIndex As Long
    XA.ReDim 1, oCust.Addresses.Count, 1, 6
    For lngIndex = 1 To oCust.Addresses.Count
        XA.Value(lngIndex, 1) = lngIndex
        XA.Value(lngIndex, 2) = oCust.Addresses(lngIndex).AddressMailing
        XA.Value(lngIndex, 3) = CreateRoleString(oCust.Addresses(lngIndex))
        XA.Value(lngIndex, 4) = oCust.Addresses(lngIndex).GetsCatalogue
        XA.Value(lngIndex, 5) = oCust.Addresses(lngIndex).Key
        XA.Value(lngIndex, 6) = oCust.Addresses(lngIndex).ForMailing
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
  '  G1.Refresh
    If XA.UpperBound(1) > 1 Then
        Me.lblRecords = XA.UpperBound(1) & " addresses"
    End If
End Sub

Private Sub G1_DblClick()
Dim frm As frmAddress
Dim lngID As Long
    Set frm = New frmAddress
    lngID = val(XA(G1.Bookmark, 5))
    frm.Component oCust.Addresses.Item(lngID)
    frm.Show vbModal
    LoadArray
End Sub
Private Sub cmdRemove_Click()
    If flgLoading Then Exit Sub
    oCust.Addresses.Remove oCust.Addresses.Item(val(XA(G1.Bookmark, 5))).Key
    LoadArray
End Sub
Private Sub cmdApproAddress_Click()
    oCust.SetApproAddressidx val(oCust.Addresses.Item(val(XA(G1.Bookmark, 5))).Key)
    LoadArray
End Sub
Private Sub cbBillTo_Click()
    oCust.SetBillToAddressidx val(oCust.Addresses.Item(val(XA(G1.Bookmark, 5))).Key)
    LoadArray
End Sub
Private Sub cbDelTo_Click()
    oCust.SetDelToAddressidx val(oCust.Addresses.Item(val(XA(G1.Bookmark, 5))).Key)
    LoadArray
End Sub
Private Sub cbOrderTo_Click()
    oCust.SetOrderToAddressidx val(oCust.Addresses.Item(val(XA(G1.Bookmark, 5))).Key)
    LoadArray
End Sub
Private Sub cmdEdit_Click()
Dim frm As frmAddress
Dim oAdd As a_Address
    If flgLoading Then Exit Sub
    Set frm = New frmAddress
    Set oAdd = oCust.Addresses.Item(XA(G1.Bookmark, 5))
    If oAdd.Addressee = "" Then oAdd.SetAddressee oCust.Title & " " & oCust.Initials & " " & oCust.Name
    frm.Component oAdd
    frm.Show vbModal
    LoadArray
    'oCust.Validate "")
End Sub
Private Function CreateRoleString(pAddress As a_Address) As String
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & vbCrLf, "")
    str = str & IIf(pAddress.DelTo = True, "Del" & vbCrLf, "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & vbCrLf, "")
    str = str & IIf(pAddress.Appro = True, "Appro" & vbCrLf, "")
    CreateRoleString = str
End Function
