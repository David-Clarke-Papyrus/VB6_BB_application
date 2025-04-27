VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPOSMain 
   BackColor       =   &H0086360B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Papyrus Point Of Sale"
   ClientHeight    =   8280
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11970
   Icon            =   "frmPOSMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLock 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Lock"
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
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   300
      Width           =   900
   End
   Begin VB.CommandButton cmdZTotal 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Cash up"
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
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   300
      Width           =   900
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
      Height          =   465
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6375
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
      TabIndex        =   17
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
         Picture         =   "frmPOSMain.frx":08CA
      End
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00D3D3CB&
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
      Width           =   900
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1755
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00404040&
      Height          =   2610
      Left            =   105
      TabIndex        =   14
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
         ItemData        =   "frmPOSMain.frx":2A4F4
         Left            =   90
         List            =   "frmPOSMain.frx":2A4FB
         TabIndex        =   15
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
         TabIndex        =   16
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
      Top             =   5985
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   90
      TabIndex        =   12
      Top             =   6795
      Width           =   11775
      Begin VB.CheckBox chkGDisc 
         BackColor       =   &H00404040&
         Height          =   255
         Left            =   5220
         TabIndex        =   25
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   8520
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5190
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5595
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
            Picture         =   "frmPOSMain.frx":2A508
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
   Begin VB.Label lblNominalDate 
      Appearance      =   0  'Flat
      BackColor       =   &H0086360B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   5550
      TabIndex        =   44
      Top             =   120
      Width           =   2220
   End
   Begin VB.Label lblSDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0086360B&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9765
      TabIndex        =   43
      Top             =   45
      Width           =   2070
   End
   Begin VB.Label lblSTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0086360B&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9765
      TabIndex        =   42
      Top             =   405
      Width           =   2070
   End
   Begin VB.Label lblOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H0086360B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   165
      TabIndex        =   18
      Top             =   405
      Width           =   2220
   End
   Begin VB.Label lblTillCode 
      Appearance      =   0  'Flat
      BackColor       =   &H0086360B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   165
      TabIndex        =   13
      Top             =   75
      Width           =   2250
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
Attribute VB_Name = "frmPOSMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enAction As eAction
Dim oCurrLine As ListItem
Dim frmCust As frmCustomer
Dim flgSaleActive As Boolean
Dim flgGDiscount As Boolean
Dim flgNewCode As Boolean
Dim flgEditItem As Boolean
Dim flgReturn As Boolean
Dim flgInvalidLine As Boolean
Dim iCurLine As Integer
Dim flgLoading As Boolean
Dim bLoggedOn As Boolean
Dim sOldStat As String
Dim sOldCode As String
Dim xPayment() As tPayment
Dim oMF As z_ManageFolders
Attribute oMF.VB_VarHelpID = -1
Dim oPOSExchange As z_POSExchange
Dim bONLINE As Boolean
Dim strOpSessionID As String
Dim strSessionID As String
Dim sBill As String
Dim sCustomer As String
Dim sPaymentType As String

'Type DestinationType
'    ForwardKey As String
'    BackwardKey As String
'    DestinationArrayIndex As Integer
'End Type
'Type StateType
'    Description As String
'    Destination(20) As DestinationType
'End Type
'Dim states(20) As StateType

Dim sLines() As tSLine




Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    ClearAllonClient
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
  Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdDelLine_Click()
    On Error GoTo errHandler
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
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdDelLine_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEditItem_Click()
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdEditItem_Click", , EA_NORERAISE
    HandleError
End Sub

Private Function LogonOperator() As Boolean
On Error GoTo errHandler
Dim bCancelled As Boolean
Dim strName As String
Dim lngStaffID As Long

   ' If Not oPC.ZSession.InSession Then
   '     oPC.ZSession.Start_Z_Session (lngStaffID)
   ' End If
    
    If oPC.ZSession.opsession.InSession Then
        oPC.ZSession.opsession.Close_OP_Session
    End If
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
         If Not oPC.ZSession.InSession Then
             oPC.ZSession.Start_Z_Session (lngStaffID)
         End If
        oPC.ZSession.opsession.START_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
        lblOperator = strName
        bLoggedOn = True
    Else
        LockAll True
        enAction = eCode
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdLogin_Click", , EA_NORERAISE
    HandleError
End Function

Private Sub cmdLogin_Click()
    LogonOperator
End Sub



Private Sub cmdOpenTill_Click()
    On Error GoTo errHandler
    MsgBox "Cash drawer is open"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdOpenTill_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRingUp_Click()
    On Error GoTo errHandler
    RingUp
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdRingUp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    On Error GoTo errHandler
    If Index = 0 Then
        ChangeLineItem False
    Else
        ChangeLineItem True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdSelect_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub


Private Sub cmdZTotal_Click()
    On Error GoTo errHandler
Dim sPass As String
Dim frm As frmSecurity
Dim lngStaffID As Long
Dim fZAct As frmZAction
Dim strName As String

    If SecurityControl(4, lngStaffID, strName, , "Enter security code to close session") Then
        If oPC.ZSession.opsession.InSession Then
            oPC.ZSession.opsession.Close_OP_Session
        End If
        If oPC.ZSession.InSession Then
            oPC.ZSession.Close_Z_Session
        End If
        Unload Me
'        oPC.ZSession.Start_Z_Session lngStaffID
'
'        lblOperator = strName
'        LockAll False
'        bLoggedOn = True
'
'        Set fZAct = New frmZAction
'        fZAct.Show vbModal
'        Set fZAct = Nothing
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdZTotal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    If Me.txtInput.Visible And Me.txtInput.Enabled Then Me.txtInput.SetFocus
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Unload Me
'        ElseIf KeyCode = vbKeyA Then
'            mnuSetup_Click
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyF1
            LoadHelp
        Case vbKeyF2
            If Not flgSaleActive And bLoggedOn Then StartSale
        
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
            ClearAllonClient
            
        Case vbKeyF12
            'OpenTill
            
        
        Case vbKeyReturn
            'handle enter key after data input
            OnEnter
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub EditItem()
    On Error GoTo errHandler
    flgEditItem = True
    Stat "Select Field with Arrow Key"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.EditItem"
End Sub

Private Sub StopEdit()
    On Error GoTo errHandler
    If Not flgEditItem Then Exit Sub
    Me.cmdEditItem.Caption = "Edit Item (F8)"
    ShowLineEdit False
    flgEditItem = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StopEdit"
End Sub

Private Sub RingUp()
    On Error GoTo errHandler
    If Not flgSaleActive Or enAction > 2 Or val(lblNumOfItems) = 0 Then Exit Sub
        
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RingUp"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_KeyUp(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    'Try to load local DB connection
    If oPC Is Nothing Then
        Set oPC = New z_POSCLIConnection
        oPC.dbConnect
    End If

    Set oPS = New z_PollingServices_Client
    Check oPS.CheckFolders, EXC_INVALIDFOLDERS, "Problem with folders"
    oPS.ConfigureFromDB
    Check oPS.TryToStartPolling, EXC_SERVERUNAVAILABLE, "Cannot poll server"
  '  oPS.LoadRecordset
    lblTillCode = oPS.TillCode
    FormatList  'Prepares columns in listview
    LoadDiscount
    cmdLogin_Click
    StandbyMode
    Me.lblNominalDate = oPC.ZSession.NominalDateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub
'Private Function Prepareconnections() As Boolean
'On Error GoTo errHandler
'Dim fInit As frmInitialize
'
'
'    Exit Function
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Initialize"
'End Function

Private Sub StandbyMode()
    On Error GoTo errHandler
    flgLoading = True
    Me.cmdEditItem.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdOpenTill.Enabled = False
    Me.cmdRingUp.Enabled = False
    
    Me.cmdLogin.Enabled = True
    Me.cmdZTotal.Enabled = True
    Me.cmdLock.Enabled = True
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StandbyMode"
End Sub

Private Sub StartSale()
    On Error GoTo errHandler
    If Not flgSaleActive Then
        ReDim xPayment(3)
        flgSaleActive = True
        lblSDate = Format(Date, "dd mmm yyyy")
        lblSTime = Format(Time, "Medium Time")
        lblPayType = "Cash" 'default
        cmdCancel.Enabled = True
        cmdLogin.Enabled = False
        cmdZTotal.Enabled = False
        cmdLock.Enabled = False
        lblCustName = "Hit 'F7' to enter"
        AddNewLine
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StartSale"
End Sub

Private Sub ClearAllonClient()
    On Error GoTo errHandler
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
    
    If Not frmCust Is Nothing Then
        Unload frmCust
        Set frmCust = Nothing
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearAllOnClient"
End Sub
Private Sub OnEnter()
    On Error GoTo errHandler
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
                ElseIf txtInput = "#" Then
                    enAction = eTitle
                    lblInput = "Title"
                    Stat "Enter Title ..."
                    If Me.chkGDisc.Value = 1 Then
                        oCurrLine.SubItems(SI_DISC) = Me.lblGDiscPercent & "%"
                    Else
                        oCurrLine.SubItems(SI_DISC) = "0%"
                    End If
                    GoTo MEX

                
                Else
                    'Try to load item from DB
                    If Not LoadProductFromCode Then
                        MsgBox "Product NOT on database!" & vbLf & _
                               "To enter Item manually type # as Item Code!", vbOKOnly, "Item not found"
                               AutoSelect txtInput
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
                If val(txtInput) = 0 Then Exit Sub
                
                lblInput = "Unit Price"
                Stat "Enter Unit Price..."
                enAction = ePrice
                flgLoading = True
                    txtInput = Format(oCurrLine.SubItems(SI_UNITPR), "0.00")
                flgLoading = False
                
            Case ePrice 'Unit Price
                If val(txtInput) = 0 Then
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
                If Not GetPayment Then
                    If Me.txtInput.Enabled And Me.txtInput.Visible Then Me.txtInput.SetFocus
                    Exit Sub
                End If
'                Me.lblPayAmount = Format(txtInput, "###0.00")
                
                ProcessSale
                Stat "Hit Enter to clear"
                enAction = eClearSale
            Case eClearSale
                ClearAllonClient
                
        End Select
    End If
    Exit Sub
MEX:
    flgLoading = True
    txtInput = ""
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.OnEnter"
End Sub


Private Function GetPayment() As Boolean
    On Error GoTo errHandler
'    If Val(txtInput) = 0 Then
'        AutoSelect txtInput
'        Exit Sub
'    End If
    If txtInput = "" Then
        flgLoading = True
        txtInput = 0
        flgLoading = False
    End If
    If Not IsNumeric(txtInput) Then
        Select Case txtInput
            Case "C" 'CreditCard
                xPayment(0).Amount = val(Me.lblTotal)
                xPayment(0).Type = "C" 'Credid Card
                Me.lblPayAmount = Me.lblTotal
                Me.lblPayType = "CARD"
            Case "K" 'Check
                xPayment(0).Amount = val(Me.lblTotal)
                xPayment(0).Type = "P" 'Paper
                Me.lblPayAmount = Me.lblTotal
                Me.lblPayType = "CHECK"
            Case "V" 'Voucher
                xPayment(0).Amount = val(Me.lblTotal)
                xPayment(0).Type = "V" 'Paper
                Me.lblPayAmount = Me.lblTotal
                Me.lblPayType = "VOUCHER"
            Case "X" 'Combination
                ShowPaymentForm
        End Select
        
    ElseIf val(txtInput) < val(lblTotal) Then
        ShowPaymentForm
    Else
        Me.lblPayAmount = Me.txtInput
        With xPayment(0)
            .Amount = val(Me.lblPayAmount)
            .Type = "M"
        End With
    End If
    
    
    
    'check if total amount input is >= to amount due
    If val(Me.lblPayAmount) < val(Me.lblTotal) Then
        'input nto valid
        flgLoading = True
        Me.txtInput = Me.lblPayAmount
        flgLoading = False
        If MsgBox("Payment amount entered is less then total amount due!" & vbLf & "Can't process sale without full due amount entered!" & vbLf & _
                  vbLf & "YES = Re-enter amount due." & vbLf & _
                  "NO = Abort sale", vbYesNo + vbExclamation, "Payment not valid!") = vbNo Then
            ClearAllonClient
        End If
        GetPayment = False
        
    Else
        'valid input
        lblChange = Format(val(Me.lblPayAmount) - val(lblTotal), "###0.00")
        lblInput = "Change"
        txtInput = Format(lblChange, "R ###0.00")
        GetPayment = True
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetPayment"
End Function

Private Sub ShowPaymentForm()
    On Error GoTo errHandler
Dim fPay As frmPayment
Dim i As Integer
Dim dTot As Double
Dim sPayType As String

        Set fPay = New frmPayment
        If txtInput = "X" Then
            xPayment(0).Amount = val(lblTotal)
        ElseIf val(txtInput) > 0 Then
            xPayment(0).Amount = val(Me.txtInput)
        End If
        fPay.Component xPayment, val(lblTotal)
        fPay.Show vbModal
        xPayment = fPay.Payment
        For i = 0 To 3
            dTot = dTot + xPayment(i).Amount
            Debug.Print Format(xPayment(i).Amount, "0.00  ") & xPayment(i).Type
            If xPayment(i).Amount > 0 Then
                If sPayType = "" Then
                    sPayType = xPayment(i).Type
                Else
                    sPayType = "X"
                End If
            End If
        Next i
        Select Case sPayType
            Case "M"
                Me.lblPayType = "CASH"
            Case "P"
                Me.lblPayType = "CHECK"
            Case "C"
                Me.lblPayType = "CARD"
            Case "V"
                Me.lblPayType = "VOUCHER"
            Case "X"
                Me.lblPayType = "MIXED"
        End Select
        Me.lblPayAmount = Format(dTot, "0.00")
        fPay.UnloadOK = True
        Unload fPay
        Set fPay = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ShowPaymentForm"
End Sub

Private Sub CalculateAll()
    On Error GoTo errHandler
Dim SubTot As Double
Dim Total As Double
Dim iQty As Integer
Dim dTotTot As Double
Dim dSubTot As Double
Dim line As ListItem
    
    '<< global discount in lblGDiscPercent is not yet inlucded in the calculation
    For Each line In Me.lstItems.ListItems
        With line
            If val(.SubItems(SI_QTY)) = 0 Or val(.Tag) = 0 Then
                GoTo nextLoop
            End If
            SubTot = val(.SubItems(SI_QTY)) * val(.Tag)
            If val(.SubItems(SI_DISC)) > 0 Then
                Total = SubTot - (SubTot * (val(.SubItems(SI_DISC)) / 100))
            Else
                Total = SubTot
            End If
            .SubItems(SI_PRICE) = Format(Total, "##0.00")
            iQty = iQty + val(.SubItems(SI_QTY))
            dSubTot = dSubTot + SubTot
            dTotTot = dTotTot + Total
        End With
nextLoop:
    Next
    Me.lblSubTotal = Format(dSubTot, "##0.00")
    Me.lblTotal = Format(dTotTot, "##0.00")
    Me.lblNumOfItems = iQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CalculateAll"
End Sub

Private Sub Stat(msg As String)
    On Error GoTo errHandler
    Status.Panels(2).Text = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Stat(msg)", msg
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
    If Button = vbRightButton Then Me.PopupMenu mnuFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_MouseUp(Button,Shift,X,Y)", Array(Button, Shift, X, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If flgSaleActive Then
        If MsgBox("There is still a sale in process!" & vbLf & _
                  "Do you want to close this Application anyway?", _
                  vbYesNo, "Sale In Process!") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    Else
        If MsgBox("Closing Papyrus POS application?", vbYesNo + vbQuestion, "Close?") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    oPC.ZSession.opsession.Close_OP_Session
    Set oMF = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub lstItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo errHandler
    With ColumnHeader
        Debug.Print .Text & " width = " & .Width
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.lstItems_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub

Private Sub lstItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.lstItems_ItemClick(Item)", Item, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuClose_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuSetup_Click()
'    On Error GoTo errHandler
''Dim fInt As New frmInitialize
'
'    With fInt
'        .Componenet oMF
'        .Show vbModal
'        .UnloadOK = True
'        Unload fInt
'        Set fInt = Nothing
'    End With
'    Me.lblTillCode = oMF.TillCode
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.mnuSetup_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub oPS_PollingStoped(msg As String)
    On Error GoTo errHandler
    If MsgBox("Automatic file transfer stopped!" & vbLf & _
               "Reason: " & msg & vbLf & vbLf & _
               "Click YES to restart it.", vbYesNo + vbExclamation, "File Transfer Stopped!") = vbYes Then
        
        oPS.StartPolling
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_PollingStoped(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub txtInput_Change()
    On Error GoTo errHandler
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
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInput_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtInput
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
Dim sTmp As String
Dim i As Integer
    Select Case enAction
        Case ePrice
            CurrencyInput Me.txtInput, KeyCode
        Case eAmPaid
            If KeyCode >= 96 And KeyCode <= 105 Then
                KeyCode = KeyCode - 48
            End If
            If Not IsNumeric(Chr(KeyCode)) And KeyCode <> 13 Then
                sTmp = UCase(Chr(KeyCode))
                flgLoading = True
                If InStr("CKVX", sTmp) > 0 Then
                    txtInput = sTmp
                Else
                    txtInput = ""
                End If
                flgLoading = False
            ElseIf KeyCode <> 13 Then
                CurrencyInput Me.txtInput, KeyCode
            End If
    
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
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
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
'    Select Case enAction
'        Case ePrice, eAmPaid
''            flgLoading = True
''                ParseCurrency txtInput
''            flgLoading = False
'    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyUp(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub


Private Sub LockAll(bLocked As Boolean)
    On Error GoTo errHandler
    Me.chkGDisc.Enabled = Not bLocked
    Me.txtInput.Enabled = Not bLocked
    Me.cmdCancel.Enabled = Not bLocked
    Me.cmdEditItem.Enabled = Not bLocked
    Me.cmdOpenTill.Enabled = Not bLocked
    Me.cmdRingUp.Enabled = Not bLocked
    Me.txtInput.Enabled = Not bLocked
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LockAll(bLocked)", bLocked
End Sub

Private Sub FormatList()
    On Error GoTo errHandler
    With Me.lstItems
        .ColumnHeaders(1).Width = 1755 'ISBN
        .ColumnHeaders(2).Width = 3540 'Title
        .ColumnHeaders(3).Width = 2730 'Author
        .ColumnHeaders(4).Width = 1110 'UnitPr.
        .ColumnHeaders(5).Width = 615  'QTY
        .ColumnHeaders(6).Width = 720  'Disc
        .ColumnHeaders(7).Width = 1035 'Price
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.FormatList"
End Sub


Private Function LoadProductFromCode() As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
    Set rs = oPC.GD.GetProduct(Trim$(Me.txtInput))
    If Not rs Is Nothing Then
        oCurrLine.SubItems(SI_TITLE) = NZS(rs!P_Title)
        oCurrLine.SubItems(SI_AUTHOR) = NZS(rs!P_MainAuthor)
        oCurrLine.SubItems(SI_UNITPR) = Format(NZ(rs!P_SAPrice), "##0.00")
        oCurrLine.Tag = NZ(rs!P_SAPrice) ' & "/" & NZ(rs!P_ID)
        oCurrLine.SubItems(SI_QTY) = "1"
        oCurrLine.SubItems(SI_PID) = NZS(rs!P_ID)
        If NZS(rs!P_Code) > "" Then
            oCurrLine.Text = NZS(rs!P_Code)
        Else
            oCurrLine.Text = NZS(rs!P_EAN)
        End If
        If Me.chkGDisc.Value = 1 Then
            oCurrLine.SubItems(SI_DISC) = Me.lblGDiscPercent & "%"
        Else
            oCurrLine.SubItems(SI_DISC) = "0%"
        End If
        CalculateAll
        LoadProductFromCode = True
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadProductFromCode"
End Function

Private Sub AddNewLine()
On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AddNewLine"
End Sub

Private Sub ParseDiscount()
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ParseDiscount"
End Sub

Private Sub ShowLineEdit(bShow As Boolean)
    On Error GoTo errHandler
    Me.cmdSelect(0).Visible = bShow
    Me.cmdSelect(1).Visible = bShow
    Me.cmdDelLine.Visible = bShow
    Me.fraList.Enabled = bShow
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ShowLineEdit(bShow)", bShow
End Sub
Private Sub ChangeLineItem(bNext As Boolean)
    On Error GoTo errHandler
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
                txtInput = val(oCurrLine.SubItems(SI_DISC))
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
            If val(txtInput) = 0 Then GoTo MEX
            CalculateAll
            If bNext Then
                'move to Discount
                enAction = eDisc
                Me.lstDisc.Visible = True
                oCurrLine.ListSubItems(SI_DISC).ForeColor = vbRed
                txtInput = val(oCurrLine.SubItems(SI_DISC))
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
            If val(txtInput) = 0 Then GoTo MEX
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ChangeLineItem(bNext)", bNext
End Sub

Private Sub ClearCurrLine()
    On Error GoTo errHandler
Dim i As Integer
    If oCurrLine Is Nothing Then Exit Sub
    If Len(oCurrLine) = 0 Then Exit Sub
    
    With oCurrLine
        .ForeColor = &H80C0FF
        For i = 1 To 5
            .ListSubItems(i).ForeColor = &H80C0FF
        Next i
   End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearCurrLine"
End Sub

Private Sub DeleteEmptyLines()
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DeleteEmptyLines"
End Sub

Private Function ItemExists() As Boolean
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ItemExists"
End Function

Private Sub LoadDiscount()
    On Error GoTo errHandler
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
    End With
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadDiscount"
End Sub

Private Sub SetDiscount()
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetDiscount"
End Sub
Private Sub SelectDiscount()
    On Error GoTo errHandler
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
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SelectDiscount"
End Sub

Private Sub LoadCustomer()
    On Error GoTo errHandler
    If Not flgSaleActive Then Exit Sub
    If frmCust Is Nothing Then Set frmCust = New frmCustomer
    With frmCust
       .Show vbModal
        If Not .bCanceled Then
            Me.lblCustName = .txtCustName
            Me.lblPayType = .sPayType
        End If
    End With
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadCustomer"
End Sub


Private Sub ProcessSale()
On Error GoTo errHandler
Dim lst As ListItem
Dim i As Integer
Dim sPID As String
Dim sExpCode As String
Dim sCustomer As String
Dim lSupervisorID As Long
Dim sPW As String
Dim iRes As Integer

    If Me.lblCustName <> "" And Left(Me.lblCustName, 8) <> "Hit 'F7'" Then
        sCustomer = Me.lblCustName
    End If
    
    '<< Here I check if any discount was assigned ---- add code to get Supervisor ID
    If GotDiscount Then
      'Just using InputBox for now...
      sPW = InputBox("Please Enter Password to allow Discount!", "Allow Discount")
      If sPW <> "admin" Then '<< Hardcoded password used! Please change....
PasswordInput:
        iRes = MsgBox("Not a valid Password!" & vbLf & "Try again?" & vbLf & _
                      "NO will remove all Discounts." & vbLf & _
                      "CANCEL will return you to Edit mode.", vbYesNoCancel + vbExclamation, "Supervisor Password")
        If iRes = vbYes Then
          GoTo PasswordInput
        ElseIf iRes = vbNo Then
          RemoveDiscount
        Else
          EditDiscount
          Exit Sub
        End If
      End If
    End If
    ReDim sLines(lstItems.ListItems.Count - 1)
    i = 0
    For Each lst In Me.lstItems.ListItems
        If lst.Text <> "" Then
        i = i + 1
'            If InStr(lst.Tag, "/") > 0 Then
'                sPID = Mid$(lst.Tag, InStr(lst.Tag, "/") + 1)
'            Else
'                sPID = ""
'            End If
     '       If InStr(lst.Tag, "N") <> 0 Then sExpCode = "P"
            With sLines(i)
                .Code = lst.SubItems(SI_AUTHOR) 'ProdCode
                .Qty = lst.SubItems(SI_QTY)
                .Author = lst.SubItems(SI_AUTHOR)
                .Title = lst.SubItems(SI_TITLE)
                .UnitPr = val(lst.SubItems(SI_UNITPR))
                .LineTot = val(lst.SubItems(SI_UNITPR)) * lst.SubItems(SI_QTY)
                .Disc = val(lst.SubItems(SI_DISC))
                .PID = Trim(lst.SubItems(SI_PID))
            End With
        End If
    Next lst
    Set oPOSExchange = New z_POSExchange
    oPOSExchange.StartPOSExchange
    oPOSExchange.SavePOSExchangeToLocal val(Me.lblTotal), val(Me.lblGDiscPercent) _
        , 0, sCustomer, 0, val(Me.lblChange), sLines, xPayment
        
    oPOSExchange.SendPOSExchange
    oPOSExchange.EndPOSExchange
    CreateBill
    '  .PaymentType = Me.lblPayType.Caption & " Payment"
    sCustomer = ""
    ReDim sLines(0)
    cmdOpenTill_Click
    frmBill.txtBill.Text = oPOSExchange.InvoiceText
    frmBill.Show vbModal
    
    Me.cmdOpenTill.Enabled = True
    Set oPOSExchange = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ProcessSale"
End Sub

Private Function GotDiscount() As Boolean
    On Error GoTo errHandler
  ' check for general discount
  If val(Me.lblGDiscPercent) > 0 Then
    GotDiscount = True
  Else ' check through the sales line for discount
    Dim lst As ListItem
    For Each lst In Me.lstItems.ListItems
      If val(lst.SubItems(SI_DISC)) > 0 Then
        GotDiscount = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GotDiscount"
End Function

Private Sub RemoveDiscount()
    On Error GoTo errHandler
Dim lst As ListItem
  Me.lblGDiscPercent = "0"
  For Each lst In Me.lstItems.ListItems
    If lst.Text <> "" Then
      lst.SubItems(SI_DISC) = "0"
    End If
  Next
  CalculateAll
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemoveDiscount"
End Sub

Private Sub EditDiscount()
    On Error GoTo errHandler
  MsgBox "To edit general discount, hit F9." & vbLf & _
        "To edit item discount, hit F8 and click on the discount field.", vbOKOnly + vbInformation, "Edit Discount?"
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.EditDiscount"
End Sub

Private Sub WriteSaleToDisc()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream
Dim sFileName As String
Dim lst As ListItem
Dim lPID As Long
Dim sExpCode As String
Dim sCustomer As String
Dim sC As String
Dim i As Integer

    sC = ","
    sFileName = "Sale" & Format(GetLastFileNum(), "00000") & ".sbk"
    Set F = fs.CreateTextFile(App.Path & "\" & sFileName)
    If Len(oPS.Guid) = 38 Then
        F.WriteLine oPS.TillCode & sC & oPS.Guid
    Else
        F.WriteLine oPS.TillCode
    End If
    If Me.lblCustName <> "" And Left(Me.lblCustName, 8) <> "Hit 'F7'" Then
        sCustomer = Me.lblCustName
    End If
    F.WriteLine Me.lblTotal & sC & oPC.GD.SalesPersonID & sC & Me.lblGDiscPercent & _
                ",0," & sCustomer
    F.WriteLine ""
    F.WriteLine "[SaleLines]"
    For Each lst In Me.lstItems.ListItems
        If lst.Text <> "" Then
            If InStr(lst.Tag, "/") > 0 Then
                lPID = val(Mid$(lst.Tag, InStr(lst.Tag, "/") + 1))
            Else
                lPID = 0
            End If
            If InStr(lst.Tag, "N") <> 0 Then sExpCode = "P"
            
            F.WriteLine lPID & sC & lst.SubItems(SI_QTY) & sC & lst.SubItems(SI_UNITPR) _
                    & sC & lst.SubItems(SI_PRICE) & ",0," & lst.SubItems(SI_DISC) _
                    & ",0,0," & lst.Text & sC & lst.SubItems(SI_AUTHOR) & sC & _
                    lst.SubItems(SI_TITLE) & sC & sExpCode
            
        End If
    Next lst
    F.WriteLine ""
    F.WriteLine "[Payment]"
    For i = 0 To UBound(xPayment)
        If i = 0 Then
            xPayment(i).Change = val(Me.lblChange)
            xPayment(i).TotPay = val(Me.lblPayAmount)
        End If
        If xPayment(i).Amount > 0 Then
            F.WriteLine xPayment(i).Amount & sC & xPayment(i).CCExpDate & sC & _
                        xPayment(i).CCNumber & sC & xPayment(i).Type & sC & xPayment(i).TotPay _
                        & sC & xPayment(i).Change
        End If
    Next i
    F.Close
    Set fs = Nothing
    Set F = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.WriteSaleToDisc"
End Sub

Private Function GetLastFileNum() As Long
    On Error GoTo errHandler
Dim lNum As Long, lTmp As Long
Dim sFile As String

    sFile = Dir(App.Path & "\*.sbk")
    Do While sFile <> ""
        lTmp = val(Mid(sFile, 5, Len(sFile) - 5))
        If lNum < lTmp Then lNum = lTmp
        sFile = Dir
    Loop
    GetLastFileNum = lNum + 1
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetLastFileNum"
End Function

Private Sub LoadHelp()
    On Error GoTo errHandler
Dim fHelp As New frmHelp
    fHelp.Show vbModal
    Set fHelp = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadHelp"
End Sub

Private Sub PrepQTY()
    On Error GoTo errHandler
    enAction = eQty
    lblInput = "Qty"
    Stat "Enter quantity ..."
    Me.txtInput.Text = "1"
    AutoSelect txtInput
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrepQTY"
End Sub

Private Sub CreateBill()
    On Error GoTo errHandler
Dim sTemp As String
Dim sT2 As String
Dim sLine As String
Dim sT As String
Dim iQty As Integer
Dim iLine As Integer
'Dim rsLine As ADODB.Recordset
Dim rsPay As ADODB.Recordset
Const iW = 20
Dim NL As String
Dim i As Integer
Dim bDisc As Integer
Dim dTotal As Double
    
    NL = vbNewLine
    
    sBill = ""
    sT = Chr(vbKeyTab)
    
    'check if we got discount
    For i = 0 To UBound(sLines)
        If sLines(i).Disc > 0 Then
            bDisc = True
            Exit For
        End If
    Next i
        
    sTemp = Centre("BOOKSHOP NAME", iW)
    
    sBill = sTemp & NL & NL
    sTemp = Centre("Tel: (021) 789 8787", iW)
    sBill = sBill & sTemp & NL
    sTemp = Centre("TAX INVOICE", iW)
    sBill = sBill & sTemp & NL
    sBill = sBill & Centre("VAT Reg. No.", iW) & NL
    sBill = sBill & Centre("0098700987", iW) & NL & NL
    
    sBill = sBill & Format(oPOSExchange.ExchangeDate, "dd mmm yyyy") & NL
    sBill = sBill & "Time:     " & Format(oPOSExchange.ExchangeDate, "hh:nn") & NL
    sBill = sBill & "Till:     " & oPC.ZSession.TillCode & NL
    sBill = sBill & "Cashier:  " & oPC.ZSession.opsession.SupervisorID & NL & NL
    If sCustomer <> "" Then
        sBill = sBill & "Customer:" & NL
        sBill = sBill & sCustomer & NL & NL
    End If
    
    sTemp = "ISBN" & NL & "Title" & NL
    sBill = sBill & sTemp
    If bDisc Then
        sTemp = "Qty " & "Unit Disc" & sT & "Total" & NL
    Else
        sTemp = "Qty " & "Unit" & sT & "Total" & NL
    End If
    
    sLine = String(iW, "-") & NL
    sBill = sBill & sTemp & sLine
   
    For i = 0 To UBound(sLines)
        With sLines(i)
            sBill = sBill & .Code & NL
            sBill = sBill & Left(.Title, iW) & NL
            If bDisc Then
                sBill = sBill & .Qty & "  " & _
                    Format(.UnitPr, "R0.00") & " %" & .Disc & sT & _
                    Format(.LineTot, "R0.00") & NL
            Else
                sBill = sBill & .Qty & "  " & _
                    Format(.UnitPr, "R0.00") & sT & _
                    Format(.LineTot, "R0.00") & NL
            End If
            sBill = sBill & sLine
            iQty = iQty + .Qty
            dTotal = dTotal + .LineTot
        End With
    Next i
    sBill = sBill & sLine
    sBill = sBill & "Total Items: " & iQty & NL
    
    sTemp = "Amount Due:"
    sT2 = Format(dTotal, "R0.00")
    sBill = sBill & sTemp & Space(iW - Len(sTemp & sT2)) & sT2 & NL
    sTemp = sPaymentType
    sT2 = Format(oPOSExchange.PaymentTotalReceived / 100, "R0.00")
    sBill = sBill & sTemp & Space(iW - Len(sTemp & sT2)) & sT2 & NL
    
    sTemp = "Change"
  '  sT2 = Format(rsPay!PAY_Change / 100, "R 0.00")
    sBill = sBill & sTemp & Space(iW - Len(sTemp & sT2)) & sT2 & NL
    
    sBill = sBill & String(iW, "=") & NL & NL
    sBill = sBill & Centre("Thank you for shopping", iW) & NL
    sBill = sBill & Centre("with us!", iW)
    
    Set rsPay = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.CreateBill"
End Sub


