VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   Caption         =   "Papyrus Point Of Sale"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11970
   Icon            =   "POSMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdZTotal 
      BackColor       =   &H00808080&
      Caption         =   "&Z Total"
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
      Height          =   405
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   300
      Width           =   1185
   End
   Begin VB.CommandButton cmdRingUp 
      BackColor       =   &H00808080&
      Caption         =   "Ring Up (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6165
      Width           =   1755
   End
   Begin VB.CommandButton cmdDelLine 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Delete Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   465
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H0080C0FF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   465
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H0080C0FF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraList 
      BorderStyle     =   0  'None
      Caption         =   "fraList"
      Height          =   3405
      Left            =   120
      TabIndex        =   22
      Top             =   765
      Width           =   11715
      Begin MSComctlLib.ListView lstItems 
         Height          =   3345
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   8438015
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ISBN"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Author"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Unit Pr."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Disc"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Price"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   0
         EndProperty
         Picture         =   "POSMain.frx":08CA
      End
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00808080&
      Caption         =   "&Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   300
      Width           =   855
   End
   Begin VB.CommandButton cmdEditItem 
      BackColor       =   &H00808080&
      Caption         =   "Edit Item (F8)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4290
      Width           =   1755
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00404040&
      Height          =   2610
      Left            =   105
      TabIndex        =   19
      Top             =   4215
      Width           =   9945
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   65.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   1590
         Left            =   1980
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   900
         Width           =   7815
      End
      Begin VB.ListBox lstDisc 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   2220
         ItemData        =   "POSMain.frx":2A4F4
         Left            =   90
         List            =   "POSMain.frx":2A4FB
         TabIndex        =   20
         Top             =   255
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblInput 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   780
         Left            =   1980
         TabIndex        =   21
         Top             =   75
         Width           =   7755
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00808080&
      Caption         =   "Close App (Ctrl+X)"
      Height          =   375
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5715
      Width           =   1755
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00404040&
      Caption         =   "Sales Time / Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   675
      Left            =   8310
      TabIndex        =   16
      Top             =   30
      Width           =   3510
      Begin VB.Label lblSDate 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label lblSTime 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   90
      TabIndex        =   13
      Top             =   6795
      Width           =   11775
      Begin VB.CheckBox chkGDisc 
         BackColor       =   &H00404040&
         Height          =   255
         Left            =   5220
         TabIndex        =   30
         Top             =   615
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   5
         X1              =   10560
         X2              =   10560
         Y1              =   120
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   4
         X1              =   8385
         X2              =   8385
         Y1              =   120
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   3
         X1              =   7080
         X2              =   7080
         Y1              =   120
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   2
         X1              =   3660
         X2              =   3660
         Y1              =   120
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   1
         X1              =   2490
         X2              =   2490
         Y1              =   105
         Y2              =   1050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   0
         X1              =   5115
         X2              =   5115
         Y1              =   120
         Y2              =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   2
         Left            =   10755
         TabIndex        =   44
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   10695
         TabIndex        =   43
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   1
         Left            =   8520
         TabIndex        =   42
         Top             =   135
         Width           =   915
      End
      Begin VB.Label lblPayAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   9330
         TabIndex        =   41
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label lblPayType 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   8520
         TabIndex        =   40
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   9555
         TabIndex        =   39
         Top             =   375
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   8670
         TabIndex        =   38
         Top             =   375
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   7350
         TabIndex        =   37
         Top             =   195
         Width           =   750
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   435
         Left            =   7170
         TabIndex        =   36
         Top             =   525
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount (F9)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   0
         Left            =   5460
         TabIndex        =   35
         Top             =   150
         Width           =   1380
      End
      Begin VB.Label lblGDiscPercent 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   6255
         TabIndex        =   34
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   6450
         TabIndex        =   33
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblGDiscType 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   5460
         TabIndex        =   32
         Top             =   570
         Width           =   705
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   5640
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   3870
         TabIndex        =   29
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   3840
         TabIndex        =   28
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   2565
         TabIndex        =   27
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lblNumOfItems 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   2715
         TabIndex        =   26
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer (F7)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblCustName 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(F7) to enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   165
         TabIndex        =   24
         Top             =   570
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00808080&
      Caption         =   "Cancel Sale (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4710
      Width           =   1755
   End
   Begin VB.CommandButton cmdOpenTill 
      BackColor       =   &H00808080&
      Caption         =   "Open Till (F12)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5265
      Width           =   1755
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   7935
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Picture         =   "POSMain.frx":2A508
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   18124
            MinWidth        =   18124
            Text            =   "To Start Sale Press F2"
            TextSave        =   "To Start Sale Press F2"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   300
      Width           =   1380
   End
   Begin VB.Label lblTillCode 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   300
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Till ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1185
      TabIndex        =   12
      Top             =   30
      Width           =   1425
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "App Setup"
      End
      Begin VB.Menu Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum eAction
    eNone = 0
    eCode = 1
    eTitle = 2
    eAuthor = 3
    eQty = 4
    eDisc = 5
    ePrice = 6
    eProceed = 7
    eAmPaid = 8
    eClearSale = 9
End Enum

Dim enAction As eAction

Dim oCurrLine As ListItem
Dim oCust As frmCustomer

Dim flgSaleActive As Boolean
Dim flgGDiscount As Boolean
Dim flgNewCode As Boolean
Dim flgEditItem As Boolean
Dim flgReturn As Boolean
Dim flgInvalidLine As Boolean

Dim iCurLine As Integer

Dim flgLoading As Boolean

Dim bLogedOn As Boolean

Dim sOldStat As String
Dim sOldCode As String

Dim xPayment() As tPayment

Dim WithEvents oEx As clsExchange
Attribute oEx.VB_VarHelpID = -1
'Dim oCode As z_ProdCode



Private Sub cmdCancel_Click()
    ClearAll
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDelLine_Click()
Dim iIndex As Integer
    If MsgBox("Are you sure you want to delete the selected item?", _
              vbYesNo + vbQuestion, "Delete Item?") = vbNo Then
        Exit Sub
    End If
    
    iIndex = oCurrLine.Index
    Set oCurrLine = Nothing
    Me.lstItems.ListItems.Remove (iIndex)
    StopEdit
    AddNewLine
    
End Sub

Private Sub cmdEditItem_Click()
    If Not flgSaleActive Then Exit Sub
    If Not flgEditItem Then
        cmdEditItem.Caption = "Stop Edit (F8)"
        Stat "Please select a line to edit..."
        ShowLineEdit True
        flgEditItem = True
        With lstItems
            If flgInvalidLine Or oCurrLine.Text = "" Then
                .ListItems.Remove (oCurrLine.Index)
                Set oCurrLine = Nothing
                flgInvalidLine = False
            End If
            .SelectedItem.Selected = False
            .SetFocus
        End With
    Else
        cmdEditItem.Caption = "Edit Item (F8)"
        flgEditItem = False
        ShowLineEdit False
        AddNewLine
    End If
End Sub



Private Sub cmdLogin_Click()
    If bLogedOn Then
        If LogOut Then
            Me.cmdLogin.Caption = "&Log In"
            Me.lblOperator = ""
            bLogedOn = False
        End If
    Else
        If LogIn Then
            Me.cmdLogin.Caption = "&Change"
            bLogedOn = True
            Stat "Press F2 to start sale..."
        End If

    End If
End Sub

Private Function LogOut() As Boolean
    LogOut = True
    Stat "Please log in..."
End Function

Private Function LogIn() As Boolean
Dim fL As New frmLogin
    With fL
        .Component 'Me.cboSalesPers.Text
        .Show vbModal
        If Not .Canceled Then
            Me.lblOperator = .SalesPerson
            LogIn = True
        Else
            Me.lblOperator = ""
            Stat "Please log in..."
            LockAll True
        End If
        Unload fL
        Set fL = Nothing
    End With
    
End Function


Private Sub cmdOpenTill_Click()
    MsgBox "Cash drawer is open"
End Sub

Private Sub cmdRingUp_Click()
    RingUp
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    If Index = 0 Then
        ChangeLineItem False
    Else
        ChangeLineItem True
    End If
End Sub

Private Sub cmdZTotal_Click()
Dim sPass As String
Dim fZAct As frmZAction
Dim fLogin As New frmLogin

    'ask for password
    With fLogin
        .Component True
        .Show vbModal
        If .Canceled Then
            Unload fLogin
            Set fLogin = Nothing
            Exit Sub
        Else
            Unload fLogin
            Set fLogin = Nothing
        End If
    End With
    
    
    Set fZAct = New frmZAction
    fZAct.Show vbModal
    
    Set fZAct = Nothing
    
End Sub

Private Sub Form_Activate()
    If Me.txtInput.Visible And Me.txtInput.Enabled Then Me.txtInput.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Unload Me
        ElseIf KeyCode = vbKeyA Then
            mnuSetup_Click
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyF1
            LoadHelp
        Case vbKeyF2
            If Not flgSaleActive And bLogedOn Then StartSale
        
        Case vbKeyF3
        
        Case vbKeyF4
        
        Case vbKeyF5
            'process sale
            If Me.cmdRingUp.Enabled Then RingUp
            
        Case vbKeyF7
            'enter customer details
            LoadCustomer
            
        Case vbKeyF8
            'Edit sale line
            cmdEditItem_Click
            
        Case vbKeyF9
            'enable, disable discount
            SetDiscount
            
        Case vbKeyF10
            'clear sale
            ClearAll
            
        Case vbKeyF12
            'OpenTill
            
        
        Case vbKeyReturn
            'handle enter key after data input
            OnEnter
    End Select
End Sub

Private Sub EditItem()
    flgEditItem = True
    Stat "Select Field with Arrow Key"
End Sub

Private Sub StopEdit()
    If Not flgEditItem Then Exit Sub
    Me.cmdEditItem.Caption = "Edit Item (F8)"
    ShowLineEdit False
    flgEditItem = False
End Sub

Private Sub RingUp()
    If Not flgSaleActive Or enAction > 2 Or Val(lblNumOfItems) = 0 Then Exit Sub
        
    Dim i As Integer
    Dim SubTot As Currency
    Dim Total As Currency
    
    
    Me.lblInput = "Amount due: " & Format(lblTotal, "R ###0.00")
    flgLoading = True
    txtInput = ""
    flgLoading = False
    Stat "Enter Cash received or 'C' = Credit Card, 'K' = Check, 'V' = Voucher, 'X' = combination payment."
    Me.cmdEditItem.Enabled = False
    Me.cmdRingUp.Enabled = False
    
    enAction = eAmPaid
    
    txtInput.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iIndex As Integer
    If flgEditItem And Not oCurrLine Is Nothing Then
        iIndex = lstItems.SelectedItem.Index
        Select Case KeyCode
            Case vbKeyPageDown
                cmdSelect_Click 0
            Case vbKeyPageUp
                cmdSelect_Click 1
            Case vbKeyUp
                If lstItems.SelectedItem.Index > 1 Then
                    ClearCurrLine
                    Set oCurrLine = lstItems.ListItems(iIndex - 1)
                    lstItems.ListItems(iIndex - 1).Selected = True
                End If
            Case vbKeyDown
                If lstItems.SelectedItem.Index < lstItems.ListItems.Count Then
                    ClearCurrLine
                    Set oCurrLine = lstItems.ListItems(iIndex + 1)
                    lstItems.ListItems(iIndex + 1).Selected = True
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    If Not Initialize Then Exit Sub
    ReDim xPayment(2)
    FormatList
    
    
    LoadDiscount
    cmdLogin_Click
    StandbyMode
End Sub

Private Function Initialize()
'Dim oEx As clsExchange
Dim fInit As frmInitialize

    'Try to load local DB connection
    Set oGD = New z_GetData
    If Not oGD.LoadDB Then
        MsgBox "Can't load local DB connection!"
        GoTo EH
    End If

    'check if we got server path and RS file
    Set oEx = New clsExchange
    If Not oEx.ServerPathOK Then
        Set fInit = New frmInitialize
        fInit.Componenet oEx
        fInit.Show vbModal
        fInit.UnloadOK = True
        If fInit.Canceled Then
            Unload fInit
            Set fInit = Nothing
            GoTo EH
        End If
        Unload fInit
        Set fInit = Nothing
    End If
    Me.lblTillCode = oEx.TillCode
    If Not oEx.RecSetLoaded Then
        MsgBox "Can't load Recordset data from disc; file missing!" & vbLf & _
               "Trying to load it from server." & vbLf & _
               "This might take a minute...", vbOKOnly, "Data file missing"
        Dim fWait As New frmWait
        fWait.Show
        If Not oEx.RequestRSFromServer Then
            Unload fWait
            Set fWait = Nothing
            MsgBox "Retreiving missing data file from server failed!" & vbLf & _
                   "Check if POS Server application is running on server and network is functional" & vbLf & _
                   "then try loading this application again." & vbLf & vbLf & _
                   "If this fails then contact:" & vbLf & _
                   "Wizards Software, Phone: (021) 426 5050", vbOKOnly, _
                   "Can't retreive missing Data file"
            
            GoTo EH
        End If
        Unload fWait
        Set fWait = Nothing
    End If
    
    
    
'    Set oCode = New z_ProdCode
    
    Initialize = True
MEX:
'    Set oEx = Nothing
    Exit Function
EH:
    Unload Me
    GoTo MEX
    
End Function

Private Sub StandbyMode()
    flgLoading = True
    Me.cmdEditItem.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdOpenTill.Enabled = False
    Me.cmdRingUp.Enabled = False
    
    Me.cmdLogin.Enabled = True
    Me.cmdZTotal.Enabled = True
    
    Me.fraList.Enabled = False
    Me.txtInput.Enabled = False
    Me.lblCustName = ""
    Me.lblPayType = ""
    Me.lblInput = ""
    Me.txtInput = ""
    Me.lblSDate = ""
    Me.lblSTime = ""
    
    If Me.cmdLogin.Enabled And Me.cmdLogin.Visible Then Me.cmdLogin.SetFocus
    flgLoading = False
End Sub

Private Sub StartSale()
    If Not flgSaleActive Then
        ReDim xPayment(3)
        flgSaleActive = True
        lblSDate = Format(Date, "dd mmm yyyy")
        lblSTime = Format(Time, "Medium Time")
        Me.lblPayType = "Cash" 'default
        Me.cmdCancel.Enabled = True
        Me.cmdLogin.Enabled = False
        Me.cmdZTotal.Enabled = False
        Me.lblCustName = "Hit 'F7' to enter"
        AddNewLine
    End If
End Sub

Private Sub ClearAll()
    If Not flgSaleActive Then Exit Sub
    If flgSaleActive And enAction <> eClearSale Then
        If MsgBox("All data of the sale in progress will be lost!" & vbLf & _
                  "Clear anyway?", vbYesNo + vbExclamation, "Clear Sale?") = vbNo Then
            Exit Sub
        End If
    End If
    If flgEditItem Then
        StopEdit
    End If
    
    If Not oCurrLine Is Nothing Then Set oCurrLine = Nothing
    
    If Not oCust Is Nothing Then
        Unload oCust
        Set oCust = Nothing
    End If
    
    Me.lstItems.ListItems.Clear
    
    flgLoading = True
    Me.txtInput = ""
    Me.lblInput = ""
    Me.lblSTime = ""
    Me.lblSDate = ""
    Me.lblCustName = ""
    Me.lblNumOfItems = "0"
    Me.lblSubTotal = "0.00"
    Me.chkGDisc.Value = 0
    Me.lblGDiscType = ""
    Me.lblGDiscPercent = 0
    Me.lblTotal = "0.00"
    Me.lblPayType = "CASH"
    Me.lblPayAmount = "0.00"
    Me.lblChange = "0.00"
    
    flgSaleActive = False
    enAction = 0
    flgGDiscount = False
    flgNewCode = False
    
    lstDisc.Visible = False
    
    Stat "To Start Press F2"
    flgLoading = False
    StandbyMode
End Sub
Private Sub OnEnter()
    If flgEditItem Then
        'In line edit mode
        If oCurrLine Is Nothing Then Exit Sub
        ChangeLineItem True
        Exit Sub
    End If
    If enAction = 0 Then
'        If Me.cboSalesPers.Text <> "" Then cmdLogin_Click
        Exit Sub
    End If
    
    If Not flgSaleActive Then
        StartSale
        Exit Sub
    End If
    If Me.lstDisc.Visible Then
        SelectDiscount
        Exit Sub
    End If
    
    If enAction > 0 Then
        Select Case enAction
                
            Case eCode 'manual Item code
                
                txtInput = Trim$(txtInput)
                If txtInput = "" Then
                    Exit Sub
                'check if an Item with identical code was allready entered
'                If ItemExists Then
'                    flgInvalidLine = True
'                    Exit Sub
'                Else
'                    flgInvalidLine = False
'                End If
'                If txtInput = "NONE" Then
'                    'txtInput = CreateNewCode
'                    txtInput = "0000X21"
'                    enAction = eTitle
'                    lblInput = "Title"
'                    Stat "Enter Title ..."
'                    GoTo MEX
'                ElseIf txtInput = "#" Then
'                    enAction = eTitle
'                    lblInput = "Title"
'                    Stat "Enter Title ..."
'                    If Me.chkGDisc.Value = 1 Then
'                        oCurrLine.SubItems(SI_DISC) = Me.lblGDiscPercent & "%"
'                    Else
'                        oCurrLine.SubItems(SI_DISC) = "0%"
'                    End If
'                    GoTo MEX

                
                Else
                    'Try to load item from DB
                    If Not LoadProductFromCode Then
                        MsgBox "Product NOT on database!" & vbLf & _
                               "To enter Item manually type # as Item Code!", vbOKOnly, "Item not found"
                               AutoSelect txtInput
'                            enAction = eTitle
'                            lblInput = "Title"
'                            Stat "Enter Title ..."
'                            If Me.chkGDisc.Value = 1 Then
'                                oCurrLine.SubItems(SI_DISC) = Me.lblGDiscPercent & "%"
'                            Else
'                                oCurrLine.SubItems(SI_DISC) = "0%"
'                            End If
'                            GoTo MEX
'                        End If
'
                        Exit Sub
                    End If
                    
                    PrepQTY

                End If
            
            
            Case eTitle 'Description
                If Len(txtInput) < 4 Then
                    AutoSelect txtInput
                    Exit Sub
                End If
                lblInput = "Author"
                Stat "Enter Author..."
                enAction = eAuthor
                GoTo MEX
            Case eAuthor
                PrepQTY
                
            Case eQty, eDisc
                If Val(txtInput) = 0 Then Exit Sub
                
                lblInput = "Unit Price"
                Stat "Enter Unit Price..."
                enAction = ePrice
                flgLoading = True
                    txtInput = Format(oCurrLine.SubItems(SI_UNITPR), "0.00")
                flgLoading = False
                
'                If Val(oCurrLine.SubItems(SI_UNITPR)) = 0 Then
'                    lblInput = "Unit Price"
'                    Stat "Enter Unit Price..."
'                    enAction = ePrice
'                    flgLoading = True
'                        txtInput = "0.00"
'                    flgLoading = False
'
'                Else
'                    CalculateAll
'                    AddNewLine
'                End If
            Case ePrice 'Unit Price
                If Val(txtInput) = 0 Then
                    AutoSelect txtInput
                    Exit Sub
                Else
                    With oCurrLine
                        If InStr(.Tag, "/") > 0 Then
                            .Tag = .SubItems(SI_UNITPR) & "N" & Mid$(.Tag, InStr(.Tag, "/"))
                        Else
                            .Tag = .SubItems(SI_UNITPR) & "N"
                        End If
                    End With
                End If
                CalculateAll
                AddNewLine
            
            Case eProceed 'Amount Received
                enAction = eAmPaid
                lblInput = "Total Amount = R " & lblTotal
                flgLoading = True
                    txtInput = "0.00"
                flgLoading = False
                
            Case eAmPaid
                GetPayment
'                Me.lblPayAmount = Format(txtInput, "###0.00")
                
                ProcessSale
                Stat "Hit Enter to clear"
                enAction = eClearSale
            Case eClearSale
                ClearAll
                
        End Select
    End If
    Exit Sub
MEX:
    flgLoading = True
    txtInput = ""
    flgLoading = False
End Sub

Private Sub GetPayment()
'    If Val(txtInput) = 0 Then
'        AutoSelect txtInput
'        Exit Sub
'    End If
    If Not IsNumeric(txtInput) Then
        Select Case txtInput
            Case "C" 'CreditCard
                xPayment(0).Amount = Val(Me.lblTotal)
                xPayment(0).Type = "C" 'Credid Card
                Me.lblPayAmount = Me.lblTotal
                Me.lblPayType = "C Card"
            Case "K" 'Check
                xPayment(0).Amount = Val(Me.lblTotal)
                xPayment(0).Type = "P" 'Paper
                Me.lblPayAmount = Me.lblTotal
                Me.lblPayType = "Check"
            Case "V" 'Voutcher
                xPayment(0).Amount = Val(Me.lblTotal)
                xPayment(0).Type = "V" 'Paper
                Me.lblPayAmount = Me.lblTotal
                Me.lblPayType = "Voutcher"
            Case "X" 'Combination
                ShowPaymentForm
                Me.lblPayType = "Comb."
        End Select
        
    ElseIf Val(txtInput) < Val(lblTotal) Then
        ShowPaymentForm
    Else
        Me.lblPayAmount = Me.txtInput
        With xPayment(0)
            .Amount = Val(Me.lblPayAmount)
            .Type = "M"
        End With
    End If
    lblChange = Format(Val(Me.lblPayAmount) - Val(lblTotal), "###0.00")
    lblInput = "Change"
    txtInput = Format(lblChange, "R ###0.00")
    
End Sub

Private Sub ShowPaymentForm()
Dim fPay As frmPayment
Dim i As Integer
Dim dTot As Double
        Set fPay = New frmPayment
        If txtInput = "X" Then
            xPayment(0).Amount = Val(lblTotal)
        ElseIf Val(txtInput) > 0 Then
            xPayment(0).Amount = Val(Me.txtInput)
        End If
        fPay.Component xPayment, Val(lblTotal)
        fPay.Show vbModal
        xPayment = fPay.Payment
        For i = 0 To 3
            dTot = dTot + xPayment(i).Amount
            Debug.Print Format(xPayment(i).Amount, "0.00  ") & xPayment(i).Type
        Next i
        Me.lblPayAmount = Format(dTot, "0.00")
        fPay.UnloadOK = True
        Unload fPay
        Set fPay = Nothing
End Sub

Private Sub CalculateAll()
Dim SubTot As Double
Dim Total As Double
Dim i As Integer
Dim iQty As Integer
Dim dTotTot As Double
Dim dSubTot As Double
Dim line As ListItem
    
    For i = 1 To Me.lstItems.ListItems.Count
        Set line = Me.lstItems.ListItems(i)
        With line
            If Val(.SubItems(SI_QTY)) = 0 Or Val(.Tag) = 0 Then
                GoTo nextLoop
            End If
            SubTot = Val(.SubItems(SI_QTY)) * Val(.Tag)
            If Val(.SubItems(SI_DISC)) > 0 Then
                Total = SubTot - (SubTot * (Val(.SubItems(SI_DISC)) / 100))
            Else
                Total = SubTot
            End If
            .SubItems(SI_PRICE) = Format(Total, "##0.00")
            iQty = iQty + Val(.SubItems(SI_QTY))
            dSubTot = dSubTot + SubTot
            dTotTot = dTotTot + Total
        End With
nextLoop:
    Next i
    Me.lblSubTotal = Format(dSubTot, "##0.00")
    Me.lblTotal = Format(dTotTot, "##0.00")
    Me.lblNumOfItems = iQty
End Sub

Private Sub Stat(msg As String)
    Status.Panels(2).Text = msg
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If flgSaleActive Then
        If MsgBox("There is still a Sale in process!" & vbLf & _
                  "Do you want to close this Application anyway?", _
                  vbYesNo, "Sale In Process!") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Set oGD = Nothing
    Set oEx = Nothing
End Sub


Private Sub lstItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ColumnHeader
        Debug.Print .Text & " width = " & .Width
    End With
End Sub

Private Sub lstItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If flgEditItem Then
        ClearCurrLine
        Set oCurrLine = Item
        flgLoading = True
            oCurrLine.ForeColor = vbRed
            Me.txtInput = oCurrLine.Text
            sOldCode = oCurrLine.Text
            Me.lblInput = "Edit Item Code..."
        flgLoading = False
        Stat "Hit F8 to exit line edit..."
        enAction = eCode
    End If
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuSetup_Click()
Dim fInt As New frmInitialize
    
    With fInt
        .Componenet oEx
        .Show vbModal
        .UnloadOK = True
        Unload fInt
        Set fInt = Nothing
    End With
    Me.lblTillCode = oEx.TillCode
    
End Sub

Private Sub oEx_PollingStoped(msg As String)
    If MsgBox("Automatic file transfer stopped!" & vbLf & _
               "Reason: " & msg & vbLf & vbLf & _
               "Click YES to restart it.", vbYesNo + vbExclamation, "File Transfer Stopped!") = vbYes Then
        
        oEx.StartPolling
    End If
End Sub

Private Sub txtInput_Change()
    Dim i As Integer
    Dim iLen As Integer
    If flgLoading Then Exit Sub
    
    txtInput.SelStart = Len(txtInput)
    Select Case enAction
        Case eCode
            oCurrLine.Text = Me.txtInput
            Me.cmdRingUp.Enabled = False
        Case eTitle
            oCurrLine.SubItems(SI_TITLE) = Me.txtInput
        Case eAuthor
            oCurrLine.SubItems(SI_AUTHOR) = Me.txtInput
        Case eQty
            oCurrLine.SubItems(SI_QTY) = Me.txtInput
        Case eDisc
            ParseDiscount
            oCurrLine.SubItems(SI_DISC) = Me.txtInput
        Case ePrice
'            ParseCurrency txtInput
            oCurrLine.SubItems(SI_UNITPR) = Me.txtInput
'            oCurrLine.Tag = Me.txtInput
        Case eAmPaid
'            ParseCurrency txtInput
    End Select
   
End Sub

Private Sub txtInput_GotFocus()
    AutoSelect txtInput
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sTmp As String
Dim i As Integer
    Select Case enAction
        Case ePrice
            CurrencyInput Me.txtInput, KeyCode
        Case eAmPaid
            If KeyCode >= 96 And KeyCode <= 105 Then
                KeyCode = KeyCode - 48
            End If
            If Not IsNumeric(Chr(KeyCode)) Then
                sTmp = UCase(Chr(KeyCode))
                flgLoading = True
                If InStr("CKVX", sTmp) > 0 Then
                    txtInput = sTmp
                Else
                    txtInput = ""
                End If
                flgLoading = False
            Else
                CurrencyInput Me.txtInput, KeyCode
            End If
    End Select
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    Select Case enAction
    
        Case eCode
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtInput) >= 13 Then KeyAscii = 0
            
        Case eQty, eDisc
            If Not IsNumeric(Chr(KeyAscii)) Then
                If KeyAscii <> vbKeyBack Then KeyAscii = 0
            End If
        Case ePrice, eAmPaid
'            If Not IsNumeric(Chr(KeyAscii)) Then
'                If Chr(KeyAscii) <> "." Then KeyAscii = 0
'                CurrencyInput Me.txtInput, KeyAscii
                KeyAscii = 0
'            End If
    End Select
    
End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)
'    Select Case enAction
'        Case ePrice, eAmPaid
''            flgLoading = True
''                ParseCurrency txtInput
''            flgLoading = False
'    End Select
End Sub


Private Sub LockAll(bLocked As Boolean)
    Me.chkGDisc.Enabled = Not bLocked
    Me.txtInput.Enabled = Not bLocked
    Me.cmdCancel.Enabled = Not bLocked
    Me.cmdEditItem.Enabled = Not bLocked
    Me.cmdOpenTill.Enabled = Not bLocked
    Me.cmdRingUp.Enabled = Not bLocked
    Me.txtInput.Enabled = Not bLocked
End Sub

Private Sub FormatList()
    With Me.lstItems
        .ColumnHeaders(1).Width = 1755 'ISBN
        .ColumnHeaders(2).Width = 3540 'Title
        .ColumnHeaders(3).Width = 2730 'Author
        .ColumnHeaders(4).Width = 1110 'UnitPr.
        .ColumnHeaders(5).Width = 615  'QTY
        .ColumnHeaders(6).Width = 720  'Disc
        .ColumnHeaders(7).Width = 1035 'Price
        .ColumnHeaders(8).Width = 0     'PID
    End With
End Sub


Private Function LoadProductFromCode() As Boolean
Dim rs As ADODB.Recordset
    Set rs = oGD.GetProductByISBN(Trim$(Me.txtInput))
    If Not rs Is Nothing Then
        oCurrLine.SubItems(SI_TITLE) = NZS(rs!P_Title)
        oCurrLine.SubItems(SI_AUTHOR) = NZS(rs!P_MainAuthor)
        oCurrLine.SubItems(SI_UNITPR) = Format(NZ(rs!P_SAPrice), "##0.00")
        oCurrLine.SubItems(SI_PID) = NZS(rs!Product_ID)
        oCurrLine.Tag = NZ(rs!P_SAPrice) & "/" & NZ(rs!Product_ID)
        oCurrLine.SubItems(SI_QTY) = "1"
        If Me.chkGDisc.Value = 1 Then
            oCurrLine.SubItems(SI_DISC) = Me.lblGDiscPercent & "%"
        Else
            oCurrLine.SubItems(SI_DISC) = "0%"
        End If
        Set rs = Nothing
        CalculateAll
        LoadProductFromCode = True
    End If
    
End Function

Private Sub AddNewLine()
Dim i As Integer
    If enAction = eDisc Then Me.lstDisc.Visible = False
    flgLoading = True
    If Me.lstItems.ListItems.Count > 0 Then
        ClearCurrLine
        DeleteEmptyLines
    End If
    Set oCurrLine = Me.lstItems.ListItems.Add
    
    For i = 1 To 6
        oCurrLine.SubItems(i) = ""
    Next i
    
    enAction = eCode
    
    Stat "Scan Barcode or enter Item Code manually... Enter # if Code not available."
    lblInput = "Item Code"
    txtInput.Text = ""
    txtInput.Enabled = True
    Me.cmdRingUp.Enabled = False
    If txtInput.Visible And txtInput.Enabled Then txtInput.SetFocus
    flgSaleActive = True
    
    flgLoading = False
    If lstItems.ListItems.Count > 1 Then
        Me.cmdEditItem.Enabled = True
        Me.cmdRingUp.Enabled = True
    End If
End Sub

Private Sub ParseDiscount()
Dim DiscOK As Boolean
Dim iOff As Integer
Dim i As Integer

    With txtInput
        If .Text = "N" Then
            .Text = "NONE"
        End If
        If IsNumeric(.Text) Then
            If Len(.Text) > 2 Then
                .Text = Left(.Text, 2)
            End If
            iOff = Len(.Text)
            For i = 0 To lstDisc.ListCount - 1
                lstDisc.Selected(i) = False
                If Left(lstDisc.List(i), iOff) = .Text Then
                    .Text = Left(lstDisc.List(i), 2)
                    .SelStart = iOff
                    .SelLength = Len(.Text) - iOff
                    lstDisc.Selected(i) = True
                    DiscOK = True
                    Exit For
                End If
            Next i
            If Not DiscOK Then
                .Text = 0
                AutoSelect txtInput
            End If
        End If
    End With
End Sub

Private Sub ShowLineEdit(bShow As Boolean)
    Me.cmdSelect(0).Visible = bShow
    Me.cmdSelect(1).Visible = bShow
    Me.cmdDelLine.Visible = bShow
    Me.fraList.Enabled = bShow
    
'    Me.lstItems.HideSelection = Not bShow
End Sub
Private Sub ChangeLineItem(bNext As Boolean)
    If Not flgEditItem Then Exit Sub
    flgLoading = True
    ClearCurrLine
    Select Case enAction
        Case eCode
            If txtInput <> sOldCode Then
                If MsgBox("The Code has been changed!" & vbLf & "To restore original Code hit YES." & vbLf & "To keep edited code hit NO.", vbYesNo, "Code Changed!") = vbYes Then
                    flgLoading = False
                    txtInput = sOldCode
                    flgLoading = True
                End If
            End If
            If bNext Then
                'move to Title
                enAction = eTitle
                oCurrLine.ListSubItems(SI_TITLE).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_TITLE)
                lblInput = "Edit Title..."
            Else
                'move to Discount
                enAction = eDisc
                Me.lstDisc.Visible = True
                oCurrLine.ListSubItems(SI_DISC).ForeColor = vbRed
                txtInput = Val(oCurrLine.SubItems(SI_DISC))
                lblInput = "Edit Discount..."
            End If
            
        Case eTitle
            If Len(txtInput) = 0 Then GoTo MEX
            If bNext Then
                'move to Author
                enAction = eAuthor
                oCurrLine.ListSubItems(SI_AUTHOR).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_AUTHOR)
                lblInput = "Edit Author..."
            Else
                'move to Code
                enAction = eCode
                oCurrLine.ForeColor = vbRed
                txtInput = oCurrLine.Text
                lblInput = "Edit Item Code..."
            End If
        Case eAuthor
            If bNext Then
                'move to Price
                enAction = ePrice
                oCurrLine.ListSubItems(SI_UNITPR).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_UNITPR)
                lblInput = "Edit Price..."
            Else
                'move to Title
                enAction = eTitle
                oCurrLine.ListSubItems(SI_TITLE).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_TITLE)
                lblInput = "Edit Title..."
            End If
        Case eQty
            If Val(txtInput) = 0 Then GoTo MEX
            CalculateAll
            If bNext Then
                'move to Discount
                enAction = eDisc
                Me.lstDisc.Visible = True
                oCurrLine.ListSubItems(SI_DISC).ForeColor = vbRed
                txtInput = Val(oCurrLine.SubItems(SI_DISC))
                lblInput = "Edit Discount..."
            Else
                'move to Price
                enAction = ePrice
                oCurrLine.ListSubItems(SI_UNITPR).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_UNITPR)
                lblInput = "Edit Price..."
                
            End If
        Case eDisc
            If Len(txtInput) = 0 Then
                flgLoading = False
                txtInput = "0%"
                flgLoading = True
            End If
            Me.lstDisc.Visible = False
            CalculateAll
            If bNext Then
                'move to Code
                enAction = eCode
                oCurrLine.ForeColor = vbRed
                txtInput = oCurrLine.Text
                lblInput = "Edit Item Code..."
            Else
                'move to Qty
                enAction = eQty
                oCurrLine.ListSubItems(SI_QTY).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_QTY)
                lblInput = "Edit Qty..."
            End If
        Case ePrice
            If Val(txtInput) = 0 Then GoTo MEX
            CalculateAll
            If bNext Then
                'move to Qty
                enAction = eQty
                oCurrLine.ListSubItems(SI_QTY).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_QTY)
                lblInput = "Edit Qty..."
                
            Else
                'move to Author
                enAction = eAuthor
                oCurrLine.ListSubItems(SI_AUTHOR).ForeColor = vbRed
                txtInput = oCurrLine.SubItems(SI_AUTHOR)
                lblInput = "Edit Author..."
            End If
    End Select
MEX:
    lstItems.Refresh
    flgLoading = False
End Sub

Private Sub ClearCurrLine()
Dim i As Integer
    If oCurrLine Is Nothing Then Exit Sub
    If Len(oCurrLine) = 0 Then Exit Sub
    
    With oCurrLine
        .ForeColor = &H80C0FF
        For i = 1 To 5
            .ListSubItems(i).ForeColor = &H80C0FF
        Next i
   End With
End Sub

Private Sub DeleteEmptyLines()
Dim i As Integer
    Set oCurrLine = Nothing
    With Me.lstItems
GoAgain:
        For i = 1 To .ListItems.Count
            If Len(.ListItems(i).Text) = 0 Then
                .ListItems.Remove (i)
                GoTo GoAgain
            End If
        Next i
    End With
End Sub

Private Function ItemExists() As Boolean
Dim lst As ListItem

    For Each lst In lstItems.ListItems
        If lst.Index <> oCurrLine.Index Then
            If lst.Text = txtInput Then
                ItemExists = True
                Set lst = Nothing
                MsgBox "Item with identical Code has allready been entered!" & vbLf & _
                       "You can enter an Item only once!" & vbLf & _
                       "You can change the Qty on the previously entered item by pressing F8!", _
                       vbOKOnly, "Item allready exists!"
                Exit For
            End If
        End If
    Next lst
End Function

Private Sub LoadDiscount()
    With lstDisc
        .Clear
        .AddItem ("10% Just Ask")
        .ItemData(.NewIndex) = 10
        .AddItem ("15% Good Customer")
        .ItemData(.NewIndex) = 15
        .AddItem ("20% Book Club")
        .ItemData(.NewIndex) = 20
        .AddItem ("30% Staff")
        .ItemData(.NewIndex) = 30
        
'        .Visible = True
    End With
        
End Sub

Private Sub SetDiscount()
Dim i As Integer
    If Not flgSaleActive Then Exit Sub
    If Me.chkGDisc.Value = 1 Then
        If MsgBox("Do you want to remove all Discount from this Sale?", vbYesNo, "Remove Dicsount?") = vbNo Then
            Exit Sub
        End If
        Me.chkGDisc.Value = 0
        For i = 1 To lstItems.ListItems.Count
            lstItems.ListItems(i).SubItems(SI_DISC) = "0%"
        Next i
        Me.lblGDiscPercent = "0"
        Me.lblGDiscType = ""
        CalculateAll
        
    Else
        Me.chkGDisc.Value = 1
        Me.txtInput.Enabled = False
        Me.lstDisc.Visible = True
        sOldStat = Status.Panels(2).Text
        Stat "Select Discount and hit Enter"
        Me.lstDisc.SetFocus
    End If
End Sub
Private Sub SelectDiscount()
Dim i As Integer
    If Not lstDisc.Visible Then Exit Sub
    For i = 1 To lstItems.ListItems.Count
        lstItems.ListItems(i).SubItems(SI_DISC) = lstDisc.ItemData(lstDisc.ListIndex) & "%"
    Next i
    Me.lblGDiscPercent = lstDisc.ItemData(lstDisc.ListIndex)
    Me.lblGDiscType = lstDisc.ListIndex + 1
    Stat sOldStat
    Me.lstDisc.Visible = False
    Me.txtInput.Enabled = True
    CalculateAll
    Me.txtInput.SetFocus
    
End Sub

Private Sub LoadCustomer()
    If Not flgSaleActive Then Exit Sub
    If oCust Is Nothing Then Set oCust = New frmCustomer
    With oCust
       .Show vbModal
        If Not .bCanceled Then
            Me.lblCustName = .txtCustName
            Me.lblPayType = .sPayType
        End If
    End With
    
End Sub


Private Sub ProcessSale()
'Dim oEx As New clsExchange
Dim lst As ListItem
Dim i As Integer
Dim strPID As String
Dim sExpCode As String
Dim sCustomer As String


    On Error GoTo EH
    If Me.lblCustName <> "" And Left(Me.lblCustName, 8) <> "Hit 'F7'" Then
        sCustomer = Me.lblCustName
    End If
    With oEx
        .AddExchange Val(Me.lblTotal), oGD.SalesPersonID, Val(Me.lblGDiscPercent) _
            , 0, sCustomer
        For Each lst In Me.lstItems.ListItems
            If lst.Text <> "" Then
               ' If InStr(lst.Tag, "/") > 0 Then
                    strPID = lst.SubItems(7)   'Val(Mid$(lst.Tag, InStr(lst.Tag, "/") + 1))
               ' Else
               '     strPID = ""
               ' End If
                If InStr(lst.Tag, "N") <> 0 Then sExpCode = "P"
                .AddSaleLine strPID, Val(lst.SubItems(SI_QTY)), Val(lst.SubItems(SI_UNITPR)), _
                        Val(lst.SubItems(SI_PRICE)), 0, Val(lst.SubItems(SI_DISC)), _
                        0, 0, lst.Text, lst.SubItems(SI_AUTHOR), _
                        lst.SubItems(SI_TITLE), sExpCode
            End If
        Next lst
        For i = 0 To UBound(xPayment)
            If i = 0 Then
                xPayment(i).Change = Val(Me.lblChange)
                xPayment(i).TotPay = Val(Me.lblPayAmount)
            End If
            If xPayment(i).Amount > 0 Then
                .AddPayment xPayment(i).Amount, xPayment(i).CCExpDate, _
                    xPayment(i).CCNumber, xPayment(i).Type, xPayment(i).TotPay, _
                    xPayment(i).Change
            End If
        Next i
        .SaveSale
        .PaymentType = Me.lblPayType & "Payment"
        .SendExchange oGD
        cmdOpenTill_Click
        frmBill.txtBill.Text = .PrintInvoice
        frmBill.Show vbModal
    End With
    
    Me.cmdOpenTill.Enabled = True
    Exit Sub
EH:
    'create a disc backup file if anything else fails
    WriteSaleToDisc
    MsgBox Err.Description
    
End Sub

Private Sub WriteSaleToDisc()
Dim fs As New FileSystemObject
Dim f As TextStream
Dim sFileName As String
Dim lst As ListItem
Dim lPID As Long
Dim sExpCode As String
Dim sCustomer As String
Dim sC As String
Dim i As Integer

    sC = ","
    sFileName = "Sale" & Format(GetLastFileNum(), "00000") & ".sbk"
    Set f = fs.CreateTextFile(App.Path & "\" & sFileName)
    If Len(oEx.Guid) = 38 Then
        f.WriteLine Me.lblTillCode & sC & oEx.Guid
    Else
        f.WriteLine Me.lblTillCode
    End If
    If Me.lblCustName <> "" And Left(Me.lblCustName, 8) <> "Hit 'F7'" Then
        sCustomer = Me.lblCustName
    End If
    f.WriteLine Me.lblTotal & sC & oGD.SalesPersonID & sC & Me.lblGDiscPercent & _
                ",0," & sCustomer
    f.WriteLine ""
    f.WriteLine "[SaleLines]"
    For Each lst In Me.lstItems.ListItems
        If lst.Text <> "" Then
            If InStr(lst.Tag, "/") > 0 Then
                lPID = Val(Mid$(lst.Tag, InStr(lst.Tag, "/") + 1))
            Else
                lPID = 0
            End If
            If InStr(lst.Tag, "N") <> 0 Then sExpCode = "P"
            
            f.WriteLine lPID & sC & lst.SubItems(SI_QTY) & sC & lst.SubItems(SI_UNITPR) _
                    & sC & lst.SubItems(SI_PRICE) & ",0," & lst.SubItems(SI_DISC) _
                    & ",0,0," & lst.Text & sC & lst.SubItems(SI_AUTHOR) & sC & _
                    lst.SubItems(SI_TITLE) & sC & sExpCode
            
        End If
    Next lst
    f.WriteLine ""
    f.WriteLine "[Payment]"
    For i = 0 To UBound(xPayment)
        If i = 0 Then
            xPayment(i).Change = Val(Me.lblChange)
            xPayment(i).TotPay = Val(Me.lblPayAmount)
        End If
        If xPayment(i).Amount > 0 Then
            f.WriteLine xPayment(i).Amount & sC & xPayment(i).CCExpDate & sC & _
                        xPayment(i).CCNumber & sC & xPayment(i).Type & sC & xPayment(i).TotPay _
                        & sC & xPayment(i).Change
        End If
    Next i
    f.Close
    Set fs = Nothing
    Set f = Nothing
End Sub

Private Function GetLastFileNum() As Long
Dim lNum As Long, lTmp As Long
Dim sFile As String

    sFile = Dir(App.Path & "\*.sbk")
    Do While sFile <> ""
        lTmp = Val(Mid(sFile, 5, Len(sFile) - 5))
        If lNum < lTmp Then lNum = lTmp
        sFile = Dir
    Loop
    GetLastFileNum = lNum + 1
End Function

Private Sub LoadHelp()
Dim fHelp As New frmHelp
    fHelp.Show vbModal
    Set fHelp = Nothing
End Sub

Private Sub PrepQTY()
    enAction = eQty
    lblInput = "Qty"
    Stat "Enter quantity ..."
    Me.txtInput.Text = "1"
    AutoSelect txtInput
End Sub
